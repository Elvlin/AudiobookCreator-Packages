import json
import os
import sys
import traceback
from pathlib import Path

PREFIX = "@@QWENJSON@@"

_model = None
_provider = "unknown"
_repo = None


def emit(obj):
    sys.stdout.write(PREFIX + json.dumps(obj, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def emit_progress(stage, message, step, total_steps):
    emit({
        "ok": True,
        "event": "progress",
        "stage": stage,
        "message": message,
        "step": int(step),
        "total_steps": int(total_steps),
    })


def _pick_device():
    pref = (os.environ.get("QWEN_WORKER_DEVICE") or "auto").strip().lower()
    try:
        import torch
        has_cuda = bool(torch.cuda.is_available())
        if pref == "cuda":
            return "cuda" if has_cuda else "cpu"
        if pref == "cpu":
            return "cpu"
        return "cuda" if has_cuda else "cpu"
    except Exception:
        return "cpu"


def _ensure_loaded():
    global _model, _provider, _repo
    if _model is not None:
        return

    repo = (os.environ.get("QWEN_WORKER_MODEL_REPO") or "Qwen/Qwen3-TTS-12Hz-1.7B-CustomVoice").strip()
    device = _pick_device()

    from qwen_tts import Qwen3TTSModel

    kwargs = {}
    if device == "cuda":
        kwargs["device_map"] = "cuda:0"
    elif device == "cpu":
        kwargs["device_map"] = "cpu"
    else:
        kwargs["device_map"] = "auto"

    _model = Qwen3TTSModel.from_pretrained(repo, **kwargs)
    _provider = device
    _repo = repo


def _synthesize(cmd):
    _ensure_loaded()

    text = (cmd.get("text") or "").strip()
    voice_path = cmd.get("voice_path") or ""
    ref_text = cmd.get("ref_text") or None
    output_path = cmd.get("output_path") or ""
    qwen = cmd.get("qwen") or {}

    if not text:
        raise RuntimeError("Input text is empty.")
    if not voice_path:
        raise RuntimeError("Voice is required.")
    if not output_path:
        raise RuntimeError("Output path is required.")

    gen_kwargs = {}
    for src, dst in (
        ("do_sample", "do_sample"),
        ("temperature", "temperature"),
        ("top_k", "top_k"),
        ("top_p", "top_p"),
        ("repetition_penalty", "repetition_penalty"),
        ("max_new_tokens", "max_new_tokens"),
    ):
        val = qwen.get(src)
        if val is not None:
            gen_kwargs[dst] = val
    x_vector_only_mode = qwen.get("x_vector_only_mode")

    use_xvec = x_vector_only_mode if x_vector_only_mode is not None else False

    # CustomVoice repo does not support generate_voice_clone(). Use generate_custom_voice()
    # with speaker id inferred from selected voice filename (stem).
    if _is_custom_voice_repo():
        speaker = _extract_custom_speaker(voice_path)
        speakers = _model.get_supported_speakers() or []
        matched_speaker = _resolve_custom_speaker_name(speaker, speakers)
        if matched_speaker is None:
            preview = ", ".join(speakers[:12]) if speakers else "(no speakers exposed by model)"
            raise RuntimeError(
                f"Selected voice '{speaker}' is not a supported CustomVoice speaker id. "
                f"Available speakers: {preview}"
            )
        audios, sr = _model.generate_custom_voice(
            text=text,
            speaker=matched_speaker,
            language="Auto",
            instruct=ref_text if ref_text else None,
            non_streaming_mode=True,
            **gen_kwargs,
        )
    else:
        if not os.path.exists(voice_path):
            raise RuntimeError(f"Voice file not found: {voice_path}")
        # Qwen upstream generate_voice_clone() trims ref_code from decoded audio using a proportional
        # estimate. That can be voice-dependent and leave residue/overflow. For ICL clone mode we run
        # the same pipeline but trim by exact decoded ref_code sample length instead.
        if use_xvec or not ref_text:
            audios, sr = _model.generate_voice_clone(
                text=text,
                ref_audio=voice_path,
                ref_text=ref_text,
                x_vector_only_mode=use_xvec,
                non_streaming_mode=True,
                **gen_kwargs,
            )
        else:
            audios, sr = _generate_voice_clone_exact_trim(
                text=text,
                ref_audio=voice_path,
                ref_text=ref_text,
                gen_kwargs=gen_kwargs,
            )

    if not audios:
        raise RuntimeError("Qwen Python backend returned no audio.")

    audio = audios[0]
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    import soundfile as sf
    sf.write(str(out_path), audio, sr)

    return {"ok": True, "sample_rate": sr, "provider": _provider, "repo": _repo}


def _is_custom_voice_repo():
    repo = (_repo or "").lower()
    return "customvoice" in repo or "custom_voice" in repo


def _resolve_custom_speaker_name(selected_name, speakers):
    if not speakers:
        return selected_name
    selected_norm = (selected_name or "").strip().lower()
    if not selected_norm:
        return None
    for s in speakers:
        if (s or "").strip().lower() == selected_norm:
            return s
    return None


def _extract_custom_speaker(voice_path):
    token_prefix = "qwen-customvoice://"
    raw = (voice_path or "").strip()
    if raw.lower().startswith(token_prefix):
        return raw[len(token_prefix):].strip()
    return Path(raw).stem.strip()


def _voice_design(cmd):
    text = (cmd.get("text") or "").strip()
    instruct = (cmd.get("instruct") or "").strip()
    language = (cmd.get("language") or "auto").strip()
    output_path = cmd.get("output_path") or ""
    qwen = cmd.get("qwen") or {}

    if not text:
        raise RuntimeError("Input text is empty.")
    if not instruct:
        raise RuntimeError("Voice design prompt is empty.")
    if not output_path:
        raise RuntimeError("Output path is required.")

    cold_start = _model is None
    total_steps = 4 if cold_start else 3
    current_step = 1

    if cold_start:
        emit_progress("model_load", "Loading VoiceDesign model...", current_step, total_steps)
        _ensure_loaded()
        current_step += 1
        emit_progress("model_loaded", "Model loaded.", current_step, total_steps)
        current_step += 1
    else:
        emit_progress("model_ready", "Model already loaded.", current_step, total_steps)
        current_step += 1

    gen_kwargs = {}
    for src, dst in (
        ("do_sample", "do_sample"),
        ("temperature", "temperature"),
        ("top_k", "top_k"),
        ("top_p", "top_p"),
        ("repetition_penalty", "repetition_penalty"),
        ("max_new_tokens", "max_new_tokens"),
    ):
        val = qwen.get(src)
        if val is not None:
            gen_kwargs[dst] = val

    emit_progress("generate_audio", "Generating voice audio...", current_step, total_steps)
    audios, sr = _model.generate_voice_design(
        text=text,
        language=language or "auto",
        instruct=instruct,
        non_streaming_mode=True,
        **gen_kwargs,
    )
    if not audios:
        raise RuntimeError("Qwen VoiceDesign returned no audio.")

    audio = audios[0]
    current_step += 1
    emit_progress("write_output", "Writing output WAV...", current_step, total_steps)
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    import soundfile as sf
    sf.write(str(out_path), audio, sr)

    return {"ok": True, "sample_rate": sr, "provider": _provider, "repo": _repo}


def _generate_voice_clone_exact_trim(text, ref_audio, ref_text, gen_kwargs):
    import torch

    prompt_items = _model.create_voice_clone_prompt(
        ref_audio=ref_audio,
        ref_text=ref_text,
        x_vector_only_mode=False,
    )
    voice_clone_prompt_dict = _model._prompt_items_to_voice_clone_prompt(prompt_items)
    ref_texts_for_ids = [it.ref_text for it in prompt_items]

    input_ids = _model._tokenize_texts([_model._build_assistant_text(text)])

    ref_ids = []
    for rt in ref_texts_for_ids:
        if rt is None or rt == "":
            ref_ids.append(None)
        else:
            ref_ids.append(_model._tokenize_texts([_model._build_ref_text(rt)])[0])

    merged_kwargs = _model._merge_generate_kwargs(**gen_kwargs)

    talker_codes_list, _ = _model.model.generate(
        input_ids=input_ids,
        ref_ids=ref_ids,
        voice_clone_prompt=voice_clone_prompt_dict,
        languages=["Auto"],
        non_streaming_mode=True,
        **merged_kwargs,
    )

    ref_code_list = voice_clone_prompt_dict.get("ref_code", None)

    codes_for_decode = []
    for i, codes in enumerate(talker_codes_list):
        if ref_code_list is not None and ref_code_list[i] is not None:
            codes_for_decode.append(torch.cat([ref_code_list[i].to(codes.device), codes], dim=0))
        else:
            codes_for_decode.append(codes)

    wavs_all, fs = _model.model.speech_tokenizer.decode([{"audio_codes": c} for c in codes_for_decode])

    wavs_out = []
    for i, wav in enumerate(wavs_all):
        if ref_code_list is None or ref_code_list[i] is None:
            wavs_out.append(wav)
            continue

        ref_code = ref_code_list[i]
        ref_wavs, _ = _model.model.speech_tokenizer.decode([{"audio_codes": ref_code}])
        exact_cut = int(ref_wavs[0].shape[0]) if ref_wavs and ref_wavs[0] is not None else 0
        if exact_cut <= 0 or exact_cut >= int(wav.shape[0]):
            # Fallback to upstream proportional trim if exact cut is unusable.
            ref_len = int(ref_code.shape[0])
            total_len = int(codes_for_decode[i].shape[0])
            exact_cut = int(ref_len / max(total_len, 1) * wav.shape[0])

        wavs_out.append(wav[exact_cut:])

    return wavs_out, fs


def main():
    try:
        emit({"ok": True, "event": "ready", "provider": _pick_device(), "repo": os.environ.get("QWEN_WORKER_MODEL_REPO", "")})
        for raw in sys.stdin:
            raw = raw.strip()
            if not raw:
                continue
            try:
                cmd = json.loads(raw)
                name = (cmd.get("cmd") or "").strip().lower()
                if name == "shutdown":
                    emit({"ok": True, "event": "bye"})
                    return 0
                if name == "synthesize":
                    emit(_synthesize(cmd))
                    continue
                if name == "voice_design":
                    emit(_voice_design(cmd))
                    continue
                emit({"ok": False, "error": f"Unknown command: {name}"})
            except Exception as ex:
                emit({"ok": False, "error": f"{type(ex).__name__}: {ex}", "trace": traceback.format_exc()})
        return 0
    except Exception as ex:
        emit({"ok": False, "error": f"{type(ex).__name__}: {ex}", "trace": traceback.format_exc()})
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
