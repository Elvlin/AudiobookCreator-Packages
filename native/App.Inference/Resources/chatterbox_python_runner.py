import argparse
import os

import torch
import torchaudio
from chatterbox.tts import ChatterboxTTS


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--text", required=True)
    parser.add_argument("--voice", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--device", default="auto")
    parser.add_argument("--repo-dir", default="")
    parser.add_argument("--exaggeration", type=float, default=0.5)
    args = parser.parse_args()

    device = args.device
    if device == "auto":
        device = "cuda" if torch.cuda.is_available() else "cpu"

    if args.repo_dir and os.path.isdir(args.repo_dir):
        model = ChatterboxTTS.from_local(args.repo_dir, device=device)
    else:
        model = ChatterboxTTS.from_pretrained(device=device)

    wav = model.generate(
        args.text,
        audio_prompt_path=args.voice,
        exaggeration=args.exaggeration,
    )
    if hasattr(wav, "detach"):
        wav = wav.detach().cpu()
    if getattr(wav, "ndim", 1) == 1:
        wav = wav.unsqueeze(0)

    torchaudio.save(args.output, wav, getattr(model, "sr", 24000))


if __name__ == "__main__":
    main()
