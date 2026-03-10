using System.Runtime.InteropServices;

namespace OrchestratorBPlugin;

/// <summary>
/// 编排器 COM 插件 - 接收或独立创建 ServiceC，提供组合操作
/// ProgID: ComPluginDemo.OrchestratorB
/// 命令: encode-hash, pipeline, info, direct
/// </summary>
[ComVisible(true)]
[Guid("E4D5F6A7-8B9C-0D1E-2F3A-B4C5D6E7F8A9")]
[ProgId("ComPluginDemo.OrchestratorB")]
[ClassInterface(ClassInterfaceType.None)]
public class OrchestratorBPlugin : IOrchestratorBPlugin
{
    private bool _initialized;

    /// <summary>宿主传递的 ServiceC 对象</summary>
    private dynamic? _serviceCObj;

    /// <summary>独立创建的 ServiceC 对象</summary>
    private dynamic? _ownServiceC;

    public string Name => "OrchestratorB";

    public string Version => "1.0.0";

    public string Description => "编排器插件 - 组合使用 ServiceC 的编码/哈希功能，支持 encode-hash, pipeline, info, direct";

    public void Initialize()
    {
        _initialized = true;
        TryCreateOwnServiceC();
    }

    public string Execute(string input)
    {
        if (!_initialized)
            Initialize();

        if (string.IsNullOrWhiteSpace(input))
            return "错误: 请输入命令，格式: <命令> [参数]\n" +
                   "可用命令:\n" +
                   "  encode-hash <文本>  - 先 Base64 编码再 SHA256 哈希\n" +
                   "  pipeline <命令1>|<命令2>|... <文本>  - 管道执行多个 ServiceC 命令\n" +
                   "  info  - 显示 ServiceC 连接状态\n" +
                   "  direct <ServiceC命令> <参数>  - 直接转发给 ServiceC 执行";

        var trimmed = input.Trim();
        var spaceIndex = trimmed.IndexOf(' ');
        var command = spaceIndex >= 0 ? trimmed[..spaceIndex].ToLowerInvariant() : trimmed.ToLowerInvariant();
        var argument = spaceIndex >= 0 ? trimmed[(spaceIndex + 1)..] : string.Empty;

        try
        {
            return command switch
            {
                "encode-hash" => EncodeHash(argument),
                "pipeline" => Pipeline(argument),
                "info" => GetInfo(),
                "direct" => Direct(argument),
                _ => $"未知命令: {command}\n可用命令: encode-hash, pipeline, info, direct"
            };
        }
        catch (Exception ex)
        {
            return $"错误: {ex.Message}";
        }
    }

    public void Shutdown()
    {
        if (_ownServiceC != null)
        {
            try
            {
                _ownServiceC.Shutdown();
                Marshal.ReleaseComObject(_ownServiceC);
            }
            catch { }
            _ownServiceC = null;
        }
        _serviceCObj = null;
        _initialized = false;
    }

    public void SetServiceCObject(object serviceCObj)
    {
        _serviceCObj = serviceCObj;
    }

    /// <summary>获取可用的 ServiceC 对象，优先使用宿主传递的</summary>
    private dynamic? GetServiceC()
    {
        return _serviceCObj ?? _ownServiceC;
    }

    /// <summary>尝试独立创建 ServiceC COM 对象</summary>
    private void TryCreateOwnServiceC()
    {
        if (_ownServiceC != null)
            return;

        try
        {
            var type = Type.GetTypeFromProgID("ComPluginDemo.ServiceC");
            if (type != null)
            {
                _ownServiceC = Activator.CreateInstance(type);
                _ownServiceC!.Initialize();
            }
        }
        catch
        {
            // ServiceC 不可用时静默失败，后续操作会报告状态
        }
    }

    /// <summary>先 Base64 编码再 SHA256 哈希</summary>
    private string EncodeHash(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "错误: 请提供要处理的文本";

        var serviceC = GetServiceC();
        if (serviceC == null)
            return "错误: ServiceC 不可用，无法执行 encode-hash";

        string base64Result = serviceC.Execute("base64enc " + input);
        string hashResult = serviceC.Execute("sha256 " + base64Result);
        return $"Base64: {base64Result}\nSHA256: {hashResult}";
    }

    /// <summary>管道执行多个 ServiceC 命令</summary>
    private string Pipeline(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "错误: 请提供管道命令，格式: <命令1>|<命令2>|... <初始文本>\n" +
                   "示例: base64enc|sha256 Hello";

        var serviceC = GetServiceC();
        if (serviceC == null)
            return "错误: ServiceC 不可用，无法执行 pipeline";

        // 解析: "cmd1|cmd2|cmd3 initialText"
        var spaceIndex = input.IndexOf(' ');
        if (spaceIndex < 0)
            return "错误: 请提供初始文本，格式: <命令1>|<命令2>|... <初始文本>";

        var pipelinePart = input[..spaceIndex];
        var currentValue = input[(spaceIndex + 1)..];
        var commands = pipelinePart.Split('|', StringSplitOptions.RemoveEmptyEntries);

        var steps = new System.Text.StringBuilder();
        steps.AppendLine($"初始值: {currentValue}");

        foreach (var cmd in commands)
        {
            var trimmedCmd = cmd.Trim();
            string result = serviceC.Execute(trimmedCmd + " " + currentValue);
            steps.AppendLine($"  [{trimmedCmd}] → {result}");
            currentValue = result;
        }

        steps.AppendLine($"最终结果: {currentValue}");
        return steps.ToString().TrimEnd();
    }

    /// <summary>显示 ServiceC 连接状态</summary>
    private string GetInfo()
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("=== OrchestratorB 状态 ===");
        sb.AppendLine($"名称: {Name}");
        sb.AppendLine($"版本: {Version}");
        sb.AppendLine();
        sb.AppendLine("--- ServiceC 连接 ---");
        sb.AppendLine($"宿主传递: {(_serviceCObj != null ? "已连接" : "未连接")}");
        sb.AppendLine($"独立创建: {(_ownServiceC != null ? "已创建" : "未创建")}");
        sb.AppendLine($"当前使用: {(_serviceCObj != null ? "宿主传递的对象" : _ownServiceC != null ? "独立创建的对象" : "不可用")}");

        var serviceC = GetServiceC();
        if (serviceC != null)
        {
            try
            {
                sb.AppendLine();
                sb.AppendLine("--- ServiceC 信息 ---");
                sb.AppendLine($"名称: {serviceC.Name}");
                sb.AppendLine($"版本: {serviceC.Version}");
                sb.AppendLine($"描述: {serviceC.Description}");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"读取 ServiceC 信息失败: {ex.Message}");
            }
        }

        return sb.ToString().TrimEnd();
    }

    /// <summary>直接转发命令给 ServiceC</summary>
    private string Direct(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "错误: 请提供 ServiceC 命令，格式: direct <命令> <参数>\n" +
                   "示例: direct base64enc Hello";

        var serviceC = GetServiceC();
        if (serviceC == null)
            return "错误: ServiceC 不可用，无法执行 direct";

        return serviceC.Execute(input);
    }
}
