using System;

namespace AudiobookCreator.UI;

internal static class UserMessageFormatter
{
    public static string FormatOperationError(string context, string? message)
    {
        var normalized = TrimToSingleLine(message);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return $"{context} failed. Try again.";
        }

        if (normalized.Contains("object is not callable", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("non-object", StringComparison.OrdinalIgnoreCase))
        {
            return "The selected model package is incompatible with the current runtime. Open Settings and re-download the active model preset.";
        }

        if (normalized.Contains("Model incomplete", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("ONNX model incomplete", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Qwen3-TTS model not downloaded", StringComparison.OrdinalIgnoreCase))
        {
            return "Model files are missing or incomplete. Open Settings and click 'Download Model Now'.";
        }

        if (normalized.Contains("InvalidProtobuf", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Protobuf parsing failed", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Qwen ONNX file is invalid or incomplete", StringComparison.OrdinalIgnoreCase))
        {
            return "The downloaded Qwen model looks corrupted. Open Settings, remove the model cache, then download it again.";
        }

        if (normalized.Contains("Qwen 1.7B ONNX set is not loadable", StringComparison.OrdinalIgnoreCase))
        {
            return "This Qwen package is not supported by the active runtime path. Switch to the supported preset and re-download the model.";
        }

        if (normalized.Contains("eSpeak-NG", StringComparison.OrdinalIgnoreCase))
        {
            return "Kitten TTS needs eSpeak-NG for phonemization. Install eSpeak-NG, then restart the app.";
        }

        if (normalized.Contains("timed out waiting for local LLM server", StringComparison.OrdinalIgnoreCase))
        {
            return "The local LLM server did not become ready. Start it again and wait for the ready status before retrying.";
        }

        if (normalized.Contains("Local server returned empty content", StringComparison.OrdinalIgnoreCase))
        {
            return "The local LLM returned no usable output. Try again, switch to a different LLM model, or simplify the request.";
        }

        if (normalized.Contains("max_tokens", StringComparison.OrdinalIgnoreCase) &&
            normalized.Contains("unsupported", StringComparison.OrdinalIgnoreCase))
        {
            return "The selected API model rejected the current token parameter. Fetch the latest model list or choose a compatible model.";
        }

        if (normalized.Contains("Repository Not Found", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("404", StringComparison.OrdinalIgnoreCase))
        {
            return "The selected download source was not found. Check the model/variant selection and try again.";
        }

        if (normalized.Contains("401", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Unauthorized", StringComparison.OrdinalIgnoreCase))
        {
            return "The request was rejected by the provider. Check the API key, account access, or gated-model permissions.";
        }

        if (normalized.Contains("403", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Forbidden", StringComparison.OrdinalIgnoreCase))
        {
            return "Access to the requested provider resource was denied. Check account permissions and model access.";
        }

        if (normalized.Contains("timed out", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("timeout", StringComparison.OrdinalIgnoreCase))
        {
            return $"{context} took too long and was canceled. Try again or reduce the workload.";
        }

        if (normalized.Contains("being used by another process", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("file in use", StringComparison.OrdinalIgnoreCase))
        {
            return "A required file is still being used by another process. Close the other app/process and try again.";
        }

        if (normalized.Contains("Python runtime check failed", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("No Python", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("python was not found", StringComparison.OrdinalIgnoreCase))
        {
            return "The Python runtime could not start. Check the configured runtime path or reinstall the bundled runtime.";
        }

        if (normalized.Contains("smoke test failed", StringComparison.OrdinalIgnoreCase))
        {
            return "The runtime started, but the verification test did not produce usable output. Reinstall the model/runtime and try again.";
        }

        if (normalized.Contains("download failed", StringComparison.OrdinalIgnoreCase))
        {
            return $"{context} download failed. Check internet access and try again.";
        }

        return normalized;
    }

    public static string TrimToSingleLine(string? text)
    {
        return string.IsNullOrWhiteSpace(text)
            ? string.Empty
            : text.Replace("\r", " ").Replace("\n", " ").Trim();
    }
}
