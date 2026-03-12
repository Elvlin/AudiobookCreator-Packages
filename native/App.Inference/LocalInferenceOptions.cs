namespace App.Inference;

public sealed class LocalInferenceOptions
{
    public string ModelCacheDir { get; init; } = "models";
    public string ModelRepoId { get; init; } = "onnx-community/chatterbox-ONNX";
    public string ModelBackend { get; init; } = string.Empty;
    public string PreferDevice { get; init; } = "auto";
    public bool ValidateOnnxRuntimeSessions { get; init; } = true;
    public int MaxNewTokens { get; init; } = 512;
    public float Exaggeration { get; init; } = 0.5f;
}
