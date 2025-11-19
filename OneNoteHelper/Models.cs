using System.Text.Json.Serialization;

namespace OneNoteHelper;

// Request/Response models for JSON-RPC communication
public class JsonRpcRequest
{
    [JsonPropertyName("method")]
    public string Method { get; set; } = "";

    [JsonPropertyName("params")]
    public Dictionary<string, object>? Params { get; set; }

    [JsonPropertyName("id")]
    public int Id { get; set; }
}

public class JsonRpcResponse
{
    [JsonPropertyName("result")]
    public object? Result { get; set; }

    [JsonPropertyName("error")]
    public JsonRpcError? Error { get; set; }

    [JsonPropertyName("id")]
    public int Id { get; set; }
}

public class JsonRpcError
{
    [JsonPropertyName("code")]
    public int Code { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = "";
}

// OneNote data models
public class NotebookInfo
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";

    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("lastModified")]
    public string LastModified { get; set; } = "";

    [JsonPropertyName("isCurrentlyViewed")]
    public bool IsCurrentlyViewed { get; set; }
}

public class ExportResult
{
    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = "";

    [JsonPropertyName("exportedPath")]
    public string ExportedPath { get; set; } = "";
}

public class VersionInfo
{
    [JsonPropertyName("version")]
    public string Version { get; set; } = "1.0.0";

    [JsonPropertyName("oneNoteInstalled")]
    public bool OneNoteInstalled { get; set; }

    [JsonPropertyName("oneNoteVersion")]
    public string OneNoteVersion { get; set; } = "";
}
