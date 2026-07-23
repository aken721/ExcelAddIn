using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace TableMagic.Cli.Mcp;

public class StdioTransport
{
    private readonly StreamReader _reader;
    private readonly StreamWriter _writer;

    private static readonly JsonSerializerOptions CachedJsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public StdioTransport()
    {
        _reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);
        _writer = new StreamWriter(Console.OpenStandardOutput(), Encoding.UTF8) { AutoFlush = true };
    }

    public async Task<string?> ReadMessageAsync()
    {
        var line = await _reader.ReadLineAsync();
        return line;
    }

    public async Task SendMessageAsync(object message)
    {
        var json = JsonSerializer.Serialize(message, CachedJsonOptions);
        await _writer.WriteLineAsync(json);
    }

    public async Task SendErrorAsync(object? id, JsonRpcError error)
    {
        var response = new JsonRpcResponse
        {
            Id = id,
            Error = error
        };
        await SendMessageAsync(response);
    }
}