import argparse

import soundfile as sf
from kittentts import KittenTTS


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--text", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--repo", required=True)
    parser.add_argument("--voice", default="expr-voice-5-m")
    parser.add_argument("--speed", type=float, default=1.0)
    args = parser.parse_args()

    model = KittenTTS(args.repo)
    audio = model.generate(args.text, voice=args.voice, speed=args.speed)
    if hasattr(audio, "detach"):
        audio = audio.detach().cpu().numpy()

    sample_rate = getattr(model, "sample_rate", 24000)
    sf.write(args.output, audio, sample_rate)


if __name__ == "__main__":
    main()
