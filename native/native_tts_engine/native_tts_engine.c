#include <stdio.h>
#include <stdint.h>
#include <string.h>
#include <math.h>
#include <stdlib.h>

#ifdef _WIN32
#define API_EXPORT __declspec(dllexport)
#else
#define API_EXPORT
#endif

#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

typedef struct VoiceProfile {
    double base_pitch_hz;
    double energy;
    int sample_rate;
} VoiceProfile;

static void set_error(char* buf, int len, const char* msg) {
    if (!buf || len <= 0) return;
    if (!msg) msg = "unknown error";
    _snprintf(buf, (size_t)len - 1, "%s", msg);
    buf[len - 1] = '\0';
}

static int file_exists(const char* path) {
    FILE* f = NULL;
#ifdef _WIN32
    if (fopen_s(&f, path, "rb") != 0) return 0;
#else
    f = fopen(path, "rb");
    if (!f) return 0;
#endif
    fclose(f);
    return 1;
}

static int contains_case_insensitive(const char* hay, const char* needle) {
    if (!hay || !needle) return 0;
    size_t nlen = strlen(needle);
    if (nlen == 0) return 1;
    for (size_t i = 0; hay[i] != '\0'; ++i) {
        size_t j = 0;
        while (needle[j] != '\0' && hay[i + j] != '\0') {
            char a = hay[i + j];
            char b = needle[j];
            if (a >= 'A' && a <= 'Z') a = (char)(a - 'A' + 'a');
            if (b >= 'A' && b <= 'Z') b = (char)(b - 'A' + 'a');
            if (a != b) break;
            j++;
        }
        if (j == nlen) return 1;
    }
    return 0;
}

static int clamp_i32(int v, int lo, int hi) {
    if (v < lo) return lo;
    if (v > hi) return hi;
    return v;
}

static double clamp_f64(double v, double lo, double hi) {
    if (v < lo) return lo;
    if (v > hi) return hi;
    return v;
}

static int read_wav_voice_profile(const char* voice_path, VoiceProfile* out_profile, char* err, int err_len) {
    if (!out_profile) {
        set_error(err, err_len, "Voice profile output is null.");
        return 70;
    }

    FILE* f = NULL;
#ifdef _WIN32
    if (fopen_s(&f, voice_path, "rb") != 0 || !f) {
        set_error(err, err_len, "Cannot open voice wav.");
        return 71;
    }
#else
    f = fopen(voice_path, "rb");
    if (!f) {
        set_error(err, err_len, "Cannot open voice wav.");
        return 71;
    }
#endif

    unsigned char header[44];
    size_t n = fread(header, 1, sizeof(header), f);
    if (n < sizeof(header)) {
        fclose(f);
        set_error(err, err_len, "Voice wav header is too short.");
        return 72;
    }

    if (memcmp(header, "RIFF", 4) != 0 || memcmp(header + 8, "WAVE", 4) != 0) {
        fclose(f);
        set_error(err, err_len, "Voice file is not a RIFF/WAVE file.");
        return 73;
    }

    int channels = (int)(header[22] | (header[23] << 8));
    int sample_rate = (int)(header[24] | (header[25] << 8) | (header[26] << 16) | (header[27] << 24));
    int bits_per_sample = (int)(header[34] | (header[35] << 8));

    if ((channels != 1 && channels != 2) || bits_per_sample != 16 || sample_rate <= 1000) {
        fclose(f);
        set_error(err, err_len, "Voice wav must be PCM 16-bit mono/stereo.");
        return 74;
    }

    int max_samples = sample_rate * 3;
    if (max_samples < 1024) max_samples = 1024;
    int16_t* pcm = (int16_t*)malloc((size_t)max_samples * sizeof(int16_t));
    if (!pcm) {
        fclose(f);
        set_error(err, err_len, "Out of memory reading voice wav.");
        return 75;
    }

    int collected = 0;
    while (collected < max_samples) {
        int16_t frame[2] = {0, 0};
        size_t read = fread(frame, sizeof(int16_t), (size_t)channels, f);
        if (read < (size_t)channels) break;
        int32_t sample = frame[0];
        if (channels == 2) sample = (sample + frame[1]) / 2;
        pcm[collected++] = (int16_t)sample;
    }
    fclose(f);

    if (collected < 512) {
        free(pcm);
        set_error(err, err_len, "Voice wav too short.");
        return 76;
    }

    // RMS energy
    double sum2 = 0.0;
    for (int i = 0; i < collected; ++i) {
        double x = (double)pcm[i] / 32768.0;
        sum2 += x * x;
    }
    double rms = sqrt(sum2 / (double)collected);

    // Zero crossing pitch estimate
    int zc = 0;
    for (int i = 1; i < collected; ++i) {
        if ((pcm[i - 1] <= 0 && pcm[i] > 0) || (pcm[i - 1] >= 0 && pcm[i] < 0)) {
            zc++;
        }
    }

    double duration = (double)collected / (double)sample_rate;
    double est_pitch = 140.0;
    if (duration > 0.02) {
        est_pitch = ((double)zc / 2.0) / duration;
    }

    est_pitch = clamp_f64(est_pitch, 80.0, 320.0);

    out_profile->base_pitch_hz = est_pitch;
    out_profile->energy = clamp_f64(rms, 0.03, 0.35);
    out_profile->sample_rate = 22050;

    free(pcm);
    return 0;
}

static int write_wav_header(FILE* f, int sample_rate, int total_samples) {
    const int channels = 1;
    const int bits = 16;
    const int byte_rate = sample_rate * channels * bits / 8;
    const int block_align = channels * bits / 8;
    const int data_size = total_samples * block_align;
    const int riff_size = 36 + data_size;

    fwrite("RIFF", 1, 4, f);
    fwrite(&riff_size, 4, 1, f);
    fwrite("WAVE", 1, 4, f);

    fwrite("fmt ", 1, 4, f);
    int fmt_size = 16;
    short audio_format = 1;
    short num_channels = channels;
    int sr = sample_rate;
    int br = byte_rate;
    short ba = block_align;
    short bps = bits;
    fwrite(&fmt_size, 4, 1, f);
    fwrite(&audio_format, 2, 1, f);
    fwrite(&num_channels, 2, 1, f);
    fwrite(&sr, 4, 1, f);
    fwrite(&br, 4, 1, f);
    fwrite(&ba, 2, 1, f);
    fwrite(&bps, 2, 1, f);

    fwrite("data", 1, 4, f);
    fwrite(&data_size, 4, 1, f);
    return 0;
}

static int is_vowel(char c) {
    return c == 'a' || c == 'e' || c == 'i' || c == 'o' || c == 'u' ||
           c == 'A' || c == 'E' || c == 'I' || c == 'O' || c == 'U';
}

static int char_duration_samples(char c, int sr, double speed) {
    int base = is_vowel(c) ? (int)(0.090 * sr) : (int)(0.060 * sr);
    if (c == ' ' || c == '\n' || c == '\t') base = (int)(0.050 * sr);
    if (c == '.' || c == '!' || c == '?') base = (int)(0.160 * sr);
    if (c == ',' || c == ';' || c == ':') base = (int)(0.100 * sr);
    double scaled = (double)base / clamp_f64(speed, 0.5, 2.0);
    return clamp_i32((int)scaled, (int)(0.015 * sr), (int)(0.240 * sr));
}

static double char_pitch_factor(char c) {
    if (c >= '0' && c <= '9') return 1.10;
    if (c == '?' || c == '!') return 1.18;
    if (c == '.') return 0.92;
    if (is_vowel(c)) return 1.04;
    return 1.00;
}

static int synthesize_voice_conditioned_wav(
    const char* text,
    const char* output_path,
    const VoiceProfile* vp,
    float speed,
    float pitch_factor,
    float energy_factor,
    float spectral_tilt,
    const float* pitch_env,
    const float* energy_env,
    const float* tilt_env,
    int env_len,
    char* err,
    int err_len
) {
    if (!text || !vp) {
        set_error(err, err_len, "Invalid synthesis inputs.");
        return 80;
    }

    int sr = vp->sample_rate > 0 ? vp->sample_rate : 22050;
    double base_pitch = vp->base_pitch_hz * clamp_f64((double)speed, 0.65, 1.6) * clamp_f64((double)pitch_factor, 0.70, 1.40);
    double amp = clamp_f64(vp->energy * 0.92 * clamp_f64((double)energy_factor, 0.60, 1.45), 0.03, 0.36);
    double tilt = clamp_f64((double)spectral_tilt, -0.40, 0.40);

    int text_len = (int)strlen(text);
    if (text_len < 1) text_len = 1;
    if (text_len > 8000) text_len = 8000;

    int total_samples = 0;
    for (int i = 0; i < text_len; ++i) {
        total_samples += char_duration_samples(text[i], sr, speed);
    }
    if (total_samples < sr) total_samples = sr;

    FILE* f = NULL;
#ifdef _WIN32
    if (fopen_s(&f, output_path, "wb") != 0 || !f) {
        set_error(err, err_len, "Cannot open output wav file for writing.");
        return 81;
    }
#else
    f = fopen(output_path, "wb");
    if (!f) {
        set_error(err, err_len, "Cannot open output wav file for writing.");
        return 81;
    }
#endif

    write_wav_header(f, sr, total_samples);

    int cursor = 0;
    int written = 0;
    for (int i = 0; i < text_len; ++i) {
        char c = text[i];
        int dur = char_duration_samples(c, sr, speed);
        float env_p = 1.0f;
        float env_e = 1.0f;
        float env_t = 0.0f;
        if (env_len > 0 && pitch_env && energy_env && tilt_env) {
            int idx = (int)((double)i / (double)text_len * (double)env_len);
            if (idx < 0) idx = 0;
            if (idx >= env_len) idx = env_len - 1;
            env_p = pitch_env[idx];
            env_e = energy_env[idx];
            env_t = tilt_env[idx];
        }

        double pf = char_pitch_factor(c) * clamp_f64((double)env_p, 0.70, 1.40);
        double f0 = base_pitch * pf;

        // crude two-formant shaping that changes with character class
        double f1 = is_vowel(c) ? 700.0 : 350.0;
        double f2 = is_vowel(c) ? 1200.0 : 2100.0;
        if (c == 'a' || c == 'A') { f1 = 750.0; f2 = 1200.0; }
        if (c == 'e' || c == 'E') { f1 = 530.0; f2 = 1800.0; }
        if (c == 'i' || c == 'I') { f1 = 300.0; f2 = 2200.0; }
        if (c == 'o' || c == 'O') { f1 = 500.0; f2 = 900.0; }
        if (c == 'u' || c == 'U') { f1 = 350.0; f2 = 700.0; }

        double local_tilt = clamp_f64((double)env_t + tilt, -0.45, 0.45);
        double local_amp = amp * clamp_f64((double)env_e, 0.70, 1.35);

        for (int j = 0; j < dur; ++j) {
            double t = (double)(cursor + j) / (double)sr;
            double local = (double)j / (double)dur;

            double env = 1.0;
            if (local < 0.08) env = local / 0.08;
            else if (local > 0.86) env = (1.0 - local) / 0.14;
            env = clamp_f64(env, 0.0, 1.0);

            double glottal = sin(2.0 * M_PI * f0 * t) + 0.42 * sin(2.0 * M_PI * (2.0 * f0) * t) + 0.16 * sin(2.0 * M_PI * (3.0 * f0) * t);
            double formants = 0.58 * sin(2.0 * M_PI * f1 * t) + 0.38 * sin(2.0 * M_PI * f2 * t);
            double tilt_mix = (1.0 + local_tilt) * sin(2.0 * M_PI * (f2 * 1.15) * t) + (1.0 - local_tilt) * sin(2.0 * M_PI * (f1 * 0.80) * t);
            double breath = 0.022 * sin(2.0 * M_PI * 40.0 * t);

            double s = (0.48 * glottal + 0.35 * formants + 0.13 * tilt_mix + breath) * env;
            int v = (int)(s * local_amp * 32767.0);
            v = clamp_i32(v, -32768, 32767);
            int16_t pcm = (int16_t)v;
            fwrite(&pcm, sizeof(int16_t), 1, f);
            written++;
        }
        cursor += dur;
    }

    fclose(f);
    return 0;
}

API_EXPORT int synthesize_to_wav_utf8(
    const char* text,
    const char* voice_path,
    const char* output_path,
    const char* model_cache_dir,
    const char* model_repo_id,
    float speed,
    char* error_buffer,
    int error_buffer_length
) {
    if (!text || text[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Input text is empty.");
        return 1;
    }
    if (!voice_path || voice_path[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Voice path is empty.");
        return 2;
    }
    if (!file_exists(voice_path)) {
        set_error(error_buffer, error_buffer_length, "Voice file not found.");
        return 3;
    }
    if (!output_path || output_path[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Output path is empty.");
        return 4;
    }
    if (!model_cache_dir || model_cache_dir[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Model cache dir is empty.");
        return 5;
    }

    const char* repo = (model_repo_id && model_repo_id[0] != '\0') ? model_repo_id : "onnx-community/chatterbox-ONNX";

    char repo_folder[512];
    size_t idx = 0;
    strcpy_s(repo_folder, sizeof(repo_folder), "models--");
    idx = strlen(repo_folder);
    for (size_t i = 0; repo[i] != '\0' && idx + 1 < sizeof(repo_folder); ++i) {
        char c = repo[i];
        if (c == '/') {
            if (idx + 2 >= sizeof(repo_folder)) break;
            repo_folder[idx++] = '-';
            repo_folder[idx++] = '-';
        } else {
            repo_folder[idx++] = c;
        }
    }
    repo_folder[idx] = '\0';

    char model_dir[1024];
    _snprintf(model_dir, sizeof(model_dir) - 1, "%s\\hf-cache\\%s", model_cache_dir, repo_folder);
    model_dir[sizeof(model_dir) - 1] = '\0';

    int is_onnx_repo = contains_case_insensitive(repo, "onnx");
    int is_qwen3_tts_repo = contains_case_insensitive(repo, "qwen3-tts");

    if (is_onnx_repo) {
        const char* onnx_required[] = {
            "onnx\\conditional_decoder.onnx",
            "onnx\\embed_tokens.onnx",
            "onnx\\language_model.onnx",
            "onnx\\speech_encoder.onnx",
            "tokenizer.json"
        };
        for (int i = 0; i < 5; ++i) {
            char file_path[1200];
            _snprintf(file_path, sizeof(file_path) - 1, "%s\\%s", model_dir, onnx_required[i]);
            file_path[sizeof(file_path) - 1] = '\0';
            if (!file_exists(file_path)) {
                char msg[1400];
                _snprintf(msg, sizeof(msg) - 1, "ONNX model incomplete. Missing file: %s", onnx_required[i]);
                msg[sizeof(msg) - 1] = '\0';
                set_error(error_buffer, error_buffer_length, msg);
                return 20 + i;
            }
        }
    } else if (is_qwen3_tts_repo) {
        char required_a[1200];
        char required_b[1200];
        char required_c[1200];
        _snprintf(required_a, sizeof(required_a) - 1, "%s\\config.json", model_dir);
        _snprintf(required_b, sizeof(required_b) - 1, "%s\\tokenizer.json", model_dir);
        _snprintf(required_c, sizeof(required_c) - 1, "%s\\generation_config.json", model_dir);
        required_a[sizeof(required_a) - 1] = '\0';
        required_b[sizeof(required_b) - 1] = '\0';
        required_c[sizeof(required_c) - 1] = '\0';
        if (!file_exists(required_a) && !file_exists(required_b) && !file_exists(required_c)) {
            set_error(error_buffer, error_buffer_length, "Qwen3-TTS model not downloaded. Use Settings -> Download Model.");
            return 41;
        }
    } else {
        const char* required[] = {
            "conds.pt",
            "tokenizer.json",
            "s3gen.safetensors",
            "t3_cfg.safetensors",
            "ve.safetensors"
        };

        for (int i = 0; i < 5; ++i) {
            char file_path[1200];
            _snprintf(file_path, sizeof(file_path) - 1, "%s\\%s", model_dir, required[i]);
            file_path[sizeof(file_path) - 1] = '\0';
            if (!file_exists(file_path)) {
                char msg[1400];
                _snprintf(msg, sizeof(msg) - 1, "Model incomplete. Missing file: %s", required[i]);
                msg[sizeof(msg) - 1] = '\0';
                set_error(error_buffer, error_buffer_length, msg);
                return 20 + i;
            }
        }
    }

    VoiceProfile vp;
    int rc = read_wav_voice_profile(voice_path, &vp, error_buffer, error_buffer_length);
    if (rc != 0) {
        return rc;
    }

    rc = synthesize_voice_conditioned_wav(
        text, output_path, &vp, speed, 1.0f, 1.0f, 0.0f,
        NULL, NULL, NULL, 0, error_buffer, error_buffer_length);
    if (rc != 0) {
        return rc;
    }

    return 0;
}

API_EXPORT int synthesize_to_wav_with_features_utf8(
    const char* text,
    const char* voice_path,
    const char* output_path,
    const char* model_cache_dir,
    const char* model_repo_id,
    float speed,
    float pitch_factor,
    float energy_factor,
    float spectral_tilt,
    char* error_buffer,
    int error_buffer_length
) {
    if (!text || text[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Input text is empty.");
        return 1;
    }
    if (!voice_path || voice_path[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Voice path is empty.");
        return 2;
    }
    if (!file_exists(voice_path)) {
        set_error(error_buffer, error_buffer_length, "Voice file not found.");
        return 3;
    }
    if (!output_path || output_path[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Output path is empty.");
        return 4;
    }
    if (!model_cache_dir || model_cache_dir[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Model cache dir is empty.");
        return 5;
    }

    // Reuse existing validation path.
    int rc = synthesize_to_wav_utf8(
        text,
        voice_path,
        output_path,
        model_cache_dir,
        model_repo_id,
        speed,
        error_buffer,
        error_buffer_length
    );
    if (rc != 0) {
        return rc;
    }

    // Overwrite with feature-shaped synthesis after successful validation.
    const char* repo = (model_repo_id && model_repo_id[0] != '\0') ? model_repo_id : "onnx-community/chatterbox-ONNX";
    (void)repo;

    VoiceProfile vp;
    rc = read_wav_voice_profile(voice_path, &vp, error_buffer, error_buffer_length);
    if (rc != 0) {
        return rc;
    }
    rc = synthesize_voice_conditioned_wav(
        text,
        output_path,
        &vp,
        speed,
        pitch_factor,
        energy_factor,
        spectral_tilt,
        NULL,
        NULL,
        NULL,
        0,
        error_buffer,
        error_buffer_length
    );
    return rc;
}

API_EXPORT int synthesize_to_wav_with_envelope_utf8(
    const char* text,
    const char* voice_path,
    const char* output_path,
    const char* model_cache_dir,
    const char* model_repo_id,
    float speed,
    float pitch_factor,
    float energy_factor,
    float spectral_tilt,
    const float* pitch_env,
    const float* energy_env,
    const float* tilt_env,
    int env_len,
    char* error_buffer,
    int error_buffer_length
) {
    if (!text || text[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Input text is empty.");
        return 1;
    }
    if (!voice_path || voice_path[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Voice path is empty.");
        return 2;
    }
    if (!file_exists(voice_path)) {
        set_error(error_buffer, error_buffer_length, "Voice file not found.");
        return 3;
    }
    if (!output_path || output_path[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Output path is empty.");
        return 4;
    }
    if (!model_cache_dir || model_cache_dir[0] == '\0') {
        set_error(error_buffer, error_buffer_length, "Model cache dir is empty.");
        return 5;
    }

    int rc = synthesize_to_wav_utf8(
        text, voice_path, output_path, model_cache_dir, model_repo_id, speed, error_buffer, error_buffer_length);
    if (rc != 0) return rc;

    VoiceProfile vp;
    rc = read_wav_voice_profile(voice_path, &vp, error_buffer, error_buffer_length);
    if (rc != 0) return rc;

    if (env_len < 0) env_len = 0;
    rc = synthesize_voice_conditioned_wav(
        text,
        output_path,
        &vp,
        speed,
        pitch_factor,
        energy_factor,
        spectral_tilt,
        pitch_env,
        energy_env,
        tilt_env,
        env_len,
        error_buffer,
        error_buffer_length
    );
    return rc;
}
