namespace App.Storage;

public sealed record ModelDownloadTelemetry(
    string RepoId,
    string FilePath,
    long BytesDownloaded,
    long? BytesTotal,
    long FileBytesDownloaded,
    long? FileBytesTotal,
    double BytesPerSecond,
    int FilesCompleted,
    int FilesTotal,
    string Message);

