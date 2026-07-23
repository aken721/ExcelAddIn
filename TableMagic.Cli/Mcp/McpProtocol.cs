using System.Text.Json.Serialization;

namespace TableMagic.Cli.Mcp;

public class JsonRpcRequest
{
    [JsonPropertyName("jsonrpc")]
    public string JsonRpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public object? Id { get; set; }

    [JsonPropertyName("method")]
    public string Method { get; set; } = "";

    [JsonPropertyName("params")]
    public object? Params { get; set; }
}

public class JsonRpcResponse
{
    [JsonPropertyName("jsonrpc")]
    public string JsonRpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public object? Id { get; set; }

    [JsonPropertyName("result")]
    public object? Result { get; set; }

    [JsonPropertyName("error")]
    public JsonRpcError? Error { get; set; }
}

public class JsonRpcError
{
    [JsonPropertyName("code")]
    public int Code { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = "";

    [JsonPropertyName("data")]
    public object? Data { get; set; }

    public static JsonRpcError ParseError => new() { Code = -32700, Message = "Parse error" };
    public static JsonRpcError InvalidRequest => new() { Code = -32600, Message = "Invalid request" };
    public static JsonRpcError MethodNotFound => new() { Code = -32601, Message = "Method not found" };
    public static JsonRpcError InvalidParams => new() { Code = -32602, Message = "Invalid params" };
    public static JsonRpcError InternalError => new() { Code = -32603, Message = "Internal error" };
}

public class InitializeParams
{
    [JsonPropertyName("protocolVersion")]
    public string ProtocolVersion { get; set; } = "";

    [JsonPropertyName("capabilities")]
    public object? Capabilities { get; set; }

    [JsonPropertyName("clientInfo")]
    public ClientInfo? ClientInfo { get; set; }
}

public class ClientInfo
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("version")]
    public string Version { get; set; } = "";
}

public class InitializeResult
{
    [JsonPropertyName("protocolVersion")]
    public string ProtocolVersion { get; set; } = "2024-11-05";

    [JsonPropertyName("capabilities")]
    public ServerCapabilities Capabilities { get; set; } = new();

    [JsonPropertyName("serverInfo")]
    public ServerInfo ServerInfo { get; set; } = new();
}

public class ServerCapabilities
{
    [JsonPropertyName("tools")]
    public ToolCapabilities? Tools { get; set; } = new();
}

public class ToolCapabilities
{
    [JsonPropertyName("listChanged")]
    public bool ListChanged { get; set; } = false;
}

public class ServerInfo
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "tablemagic";

    [JsonPropertyName("version")]
    public string Version { get; set; } = "1.0.0";
}

public class ToolsListParams
{
}

public class ToolsListResult
{
    [JsonPropertyName("tools")]
    public List<McpTool> Tools { get; set; } = new();
}

public class McpTool
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("description")]
    public string Description { get; set; } = "";

    [JsonPropertyName("inputSchema")]
    public object? InputSchema { get; set; }
}

public class ToolCallParams
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("arguments")]
    public Dictionary<string, object>? Arguments { get; set; }
}

public class ToolCallResult
{
    [JsonPropertyName("content")]
    public List<ToolCallContent> Content { get; set; } = new();

    [JsonPropertyName("isError")]
    public bool IsError { get; set; } = false;
}

public class ToolCallContent
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "text";

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";
}