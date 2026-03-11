# Audiobook Creator Lite (GitHub Edition)

This is the **Lite** public source version of Audiobook Creator.

## Lite Version Scope

- Supported model: **Kitten TTS Mini 0.8** (`KittenML/kitten-tts-mini-0.8`)
- Local native ONNX inference (no Qwen / Chatterbox in this repo copy)
- Audiobook-oriented text chunking + punctuation-aware pauses
- Basic model download from the UI (Kitten model files)

## Removed From This GitHub Lite Version

- Qwen 3 support
- Chatterbox support
- Voice Cloning tab and voice cloning features
- Qwen Python worker runtime and related code

## Main / Private Version

The main version supports:

- **Qwen 3 TTS**
- **Chatterbox**
- additional advanced features not included here

## Requirements

- Windows 10/11
- .NET SDK 7.0 (for building from source)
- NVIDIA GPU optional (CPU also works)
- **eSpeak-NG** (recommended / required by Kitten backend for good phonemization)

Install eSpeak-NG:
- UI button opens installer URL, or install manually from the official project release page.

## Build

From repo root (`github` folder):

```powershell
cd native
 dotnet build App.UI\App.UI.csproj -c Release
```

The build outputs to:

- `customer_release\` (inside this `github` folder)

## Run

```powershell
.\customer_release\AudiobookCreator.exe
```

## Notes

- This repo copy is intentionally simplified for GitHub publishing.
- The app UI includes a clear Lite-version notice so users know Qwen/Chatterbox are not included here.