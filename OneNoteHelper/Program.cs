using System.Text.Json;
using OneNoteHelper;

class Program
{
    static void Main(string[] args)
    {
        // JSON-RPC over stdin/stdout
        // This allows Go to communicate with this helper program

        try
        {
            // Read JSON request from stdin
            var input = Console.In.ReadToEnd();

            if (string.IsNullOrWhiteSpace(input))
            {
                WriteError(-32600, "Invalid Request: Empty input", 0);
                return;
            }

            JsonRpcRequest? request;
            try
            {
                request = JsonSerializer.Deserialize<JsonRpcRequest>(input);
            }
            catch (JsonException ex)
            {
                WriteError(-32700, "Parse error: " + ex.Message, 0);
                return;
            }

            if (request == null)
            {
                WriteError(-32600, "Invalid Request: Null request", 0);
                return;
            }

            // Process the request
            ProcessRequest(request);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Fatal error: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
            WriteError(-32603, "Internal error: " + ex.Message, 0);
        }
    }

    static void ProcessRequest(JsonRpcRequest request)
    {
        try
        {
            using var service = new OneNoteService();
            object? result = null;

            switch (request.Method)
            {
                case "GetVersion":
                    result = service.GetVersionInfo();
                    break;

                case "GetNotebooks":
                    result = service.GetNotebooks();
                    break;

                case "ExportNotebook":
                    if (request.Params == null ||
                        !request.Params.ContainsKey("notebookId") ||
                        !request.Params.ContainsKey("destinationPath"))
                    {
                        WriteError(-32602, "Invalid params: notebookId and destinationPath required", request.Id);
                        return;
                    }

                    var notebookId = request.Params["notebookId"].ToString() ?? "";
                    var destPath = request.Params["destinationPath"].ToString() ?? "";
                    var format = request.Params.ContainsKey("format")
                        ? request.Params["format"].ToString() ?? "onepkg"
                        : "onepkg";

                    result = service.ExportNotebook(notebookId, destPath, format);
                    break;

                case "ExportAllNotebooks":
                    if (request.Params == null || !request.Params.ContainsKey("destinationPath"))
                    {
                        WriteError(-32602, "Invalid params: destinationPath required", request.Id);
                        return;
                    }

                    var destPathAll = request.Params["destinationPath"].ToString() ?? "";
                    var formatAll = request.Params.ContainsKey("format")
                        ? request.Params["format"].ToString() ?? "onepkg"
                        : "onepkg";
                    result = service.ExportAllNotebooks(destPathAll, formatAll);
                    break;

                default:
                    WriteError(-32601, $"Method not found: {request.Method}", request.Id);
                    return;
            }

            // Write successful response
            var response = new JsonRpcResponse
            {
                Result = result,
                Id = request.Id
            };

            var json = JsonSerializer.Serialize(response, new JsonSerializerOptions
            {
                WriteIndented = false
            });
            Console.WriteLine(json);
        }
        catch (InvalidOperationException ex)
        {
            WriteError(-32000, ex.Message, request.Id);
        }
        catch (Exception ex)
        {
            WriteError(-32603, "Internal error: " + ex.Message, request.Id);
        }
    }

    static void WriteError(int code, string message, int id)
    {
        var response = new JsonRpcResponse
        {
            Error = new JsonRpcError
            {
                Code = code,
                Message = message
            },
            Id = id
        };

        var json = JsonSerializer.Serialize(response, new JsonSerializerOptions
        {
            WriteIndented = false
        });
        Console.WriteLine(json);
    }
}
