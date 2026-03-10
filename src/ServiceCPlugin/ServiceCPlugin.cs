using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace ServiceCPlugin;

/// <summary>
/// 编码/哈希工具 COM 插件
/// ProgID: ComPluginDemo.ServiceC
/// 命令: base64enc, base64dec, urlenc, urldec, sha256, md5
/// </summary>
[ComVisible(true)]
[Guid("C2B3D4E5-6F7A-8B9C-0D1E-F2A3B4C5D6E7")]
[ProgId("ComPluginDemo.ServiceC")]
[ClassInterface(ClassInterfaceType.None)]
public class ServiceCPlugin : IServiceCPlugin
{
    private bool _initialized;

    public string Name => "ServiceC";

    public string Version => "1.0.0";

    public string Description => "编码/哈希工具服务 - 支持 base64, url编码, sha256, md5";

    public void Initialize()
    {
        _initialized = true;
    }

    public string Execute(string input)
    {
        if (!_initialized)
            Initialize();

        if (string.IsNullOrWhiteSpace(input))
            return "错误: 请输入命令和参数，格式: <命令> <参数>\n" +
                   "可用命令: base64enc, base64dec, urlenc, urldec, sha256, md5";

        var trimmed = input.Trim();
        var spaceIndex = trimmed.IndexOf(' ');
        var command = spaceIndex >= 0 ? trimmed[..spaceIndex].ToLowerInvariant() : trimmed.ToLowerInvariant();
        var argument = spaceIndex >= 0 ? trimmed[(spaceIndex + 1)..] : string.Empty;

        try
        {
            return command switch
            {
                "base64enc" => Base64Encode(argument),
                "base64dec" => Base64Decode(argument),
                "urlenc" => UrlEncode(argument),
                "urldec" => UrlDecode(argument),
                "sha256" => Sha256Hash(argument),
                "md5" => Md5Hash(argument),
                _ => $"未知命令: {command}\n可用命令: base64enc, base64dec, urlenc, urldec, sha256, md5"
            };
        }
        catch (Exception ex)
        {
            return $"错误: {ex.Message}";
        }
    }

    public void Shutdown()
    {
        _initialized = false;
    }

    private static string Base64Encode(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "错误: 请提供要编码的文本";
        var bytes = Encoding.UTF8.GetBytes(input);
        return Convert.ToBase64String(bytes);
    }

    private static string Base64Decode(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "错误: 请提供要解码的 Base64 字符串";
        var bytes = Convert.FromBase64String(input);
        return Encoding.UTF8.GetString(bytes);
    }

    private static string UrlEncode(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "错误: 请提供要编码的文本";
        return Uri.EscapeDataString(input);
    }

    private static string UrlDecode(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "错误: 请提供要解码的 URL 编码字符串";
        return Uri.UnescapeDataString(input);
    }

    private static string Sha256Hash(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "错误: 请提供要哈希的文本";
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(input));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    private static string Md5Hash(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "错误: 请提供要哈希的文本";
        var bytes = MD5.HashData(Encoding.UTF8.GetBytes(input));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }
}
