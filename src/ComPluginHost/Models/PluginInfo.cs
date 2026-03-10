using System;

namespace ComPluginHost.Models;

/// <summary>
/// 已加载的 COM 插件信息模型
/// </summary>
public class PluginInfo
{
    public string ProgId { get; set; } = string.Empty;
    public string ManifestPath { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Version { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public object? ComObject { get; set; }
    public bool IsLoaded { get; set; }
    public string? ErrorMessage { get; set; }

    /// <summary>加载来源: "Manifest" 或 "注册表" 或 "注册表 (manifest失败回退)"</summary>
    public string LoadSource { get; set; } = "Manifest";

    /// <summary>动态激活上下文句柄 (仅清单动态加载时使用)</summary>
    public IntPtr ActCtxHandle { get; set; }

    /// <summary>激活上下文 Cookie (用于 DeactivateActCtx)</summary>
    public IntPtr ActCtxCookie { get; set; }

    public string DisplayName => IsLoaded ? $"{Name} v{Version}" : $"[失败] {ProgId}";
}
