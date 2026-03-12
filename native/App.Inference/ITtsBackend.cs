namespace App.Inference;

public interface ITtsBackend
{
    string Name { get; }
    Task SynthesizeAsync(TtsRequest request, CancellationToken ct = default);
}
