using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using ComPluginHost.Models;

namespace ComPluginHost.Services;

/// <summary>
/// COM 插件加载器 - 支持免注册 COM (Reg-Free COM)
///
/// 加载流程:
///   1. 优先使用 exe.manifest 中声明的 SxS 依赖 (静态激活上下文)
///   2. 若失败，使用 Activation Context API 从插件独立清单动态创建激活上下文
///   3. 在激活上下文中调用 Type.GetTypeFromProgID + Activator.CreateInstance
/// </summary>
public class ComPluginLoader
{
    #region Win32 Activation Context API

    private static readonly IntPtr INVALID_HANDLE_VALUE = new(-1);

    private const uint ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID = 0x004;

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateActCtx(ref ACTCTX pActCtx);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ActivateActCtx(IntPtr hActCtx, out IntPtr lpCookie);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeactivateActCtx(uint dwFlags, IntPtr lpCookie);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern void ReleaseActCtx(IntPtr hActCtx);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ACTCTX
    {
        public int cbSize;
        public uint dwFlags;
        public string lpSource;
        public ushort wProcessorArchitecture;
        public ushort wLangId;
        public string? lpAssemblyDirectory;
        public string? lpResourceName;
        public string? lpApplicationName;
        public IntPtr hModule;
    }

    #endregion

    /// <summary>
    /// 从配置文件加载插件配置列表
    /// </summary>
    public static List<PluginEntry> LoadPluginConfig(string configPath)
    {
        if (!File.Exists(configPath))
            return new List<PluginEntry>();

        var json = File.ReadAllText(configPath);
        var config = JsonSerializer.Deserialize<PluginConfig>(json,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        return config?.Plugins ?? new List<PluginEntry>();
    }

    /// <summary>
    /// 保存插件配置到文件
    /// </summary>
    public static void SavePluginConfig(string configPath, List<PluginEntry> entries)
    {
        var config = new PluginConfig { Plugins = entries };
        var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(configPath, json);
    }

    /// <summary>
    /// 加载单个 COM 插件
    ///
    /// 策略 (按优先级):
    ///   1. 若配置了 ManifestPath 且清单文件存在 → Activation Context API 免注册加载
    ///   2. 回退到注册表 → Type.GetTypeFromProgID (支持 VB6 IDE 调试)
    ///
    /// 调试 VB6 时: 删除输出目录中的 manifest 文件, 加载器自动回退到注册表,
    /// 连接到 VB6 IDE 正在运行的实例, 即可命中断点.
    /// </summary>
    public static PluginInfo LoadPlugin(PluginEntry entry, string baseDir)
    {
        var info = new PluginInfo
        {
            ProgId = entry.ProgId,
            ManifestPath = entry.ManifestPath
        };

        string manifestError = "";

        // 策略 1: 优先通过 manifest 免注册加载
        if (!string.IsNullOrEmpty(entry.ManifestPath))
        {
            var manifestFullPath = Path.IsPathRooted(entry.ManifestPath)
                ? entry.ManifestPath
                : Path.GetFullPath(Path.Combine(baseDir, entry.ManifestPath));

            if (File.Exists(manifestFullPath))
            {
                var result = LoadPluginWithActivationContext(info, manifestFullPath);
                if (result.IsLoaded)
                    return result;
                // manifest 加载失败, 记录原因, 继续尝试注册表
                manifestError = result.ErrorMessage ?? "";
            }
            else
            {
                manifestError = $"清单文件不存在: {manifestFullPath}";
            }
        }

        // 策略 2: 回退到注册表 (也支持 VB6 IDE 调试模式)
        string registryError;
        var type = Type.GetTypeFromProgID(entry.ProgId);
        if (type != null)
        {
            var result = CreatePluginFromType(info, type);
            if (result.IsLoaded)
            {
                result.LoadSource = string.IsNullOrEmpty(manifestError) ? "注册表" : "注册表 (manifest失败回退)";
                return result;
            }
            registryError = result.ErrorMessage ?? "创建 COM 对象失败";
        }
        else
        {
            registryError = $"找不到 ProgID: {entry.ProgId}";
        }

        // 两种策略都失败
        info.ErrorMessage = "所有加载策略均失败:\n";
        if (!string.IsNullOrEmpty(manifestError))
            info.ErrorMessage += $"  [Manifest] {manifestError}\n";
        info.ErrorMessage += $"  [注册表] {registryError}\n" +
                            "请确认:\n" +
                            "  1. 组件清单和 DLL 在正确位置 (免注册)\n" +
                            "  2. 或 COM 已注册 / VB6 IDE 正在运行 (注册表)";
        return info;
    }

    /// <summary>
    /// 使用 Activation Context API 从独立清单加载 COM 插件
    /// </summary>
    private static PluginInfo LoadPluginWithActivationContext(PluginInfo info, string manifestPath)
    {
        var assemblyDir = Path.GetDirectoryName(Path.GetFullPath(manifestPath)) ?? "";

        var actCtx = new ACTCTX
        {
            cbSize = Marshal.SizeOf<ACTCTX>(),
            dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID,
            lpSource = manifestPath,
            lpAssemblyDirectory = assemblyDir
        };

        IntPtr hActCtx = CreateActCtx(ref actCtx);
        if (hActCtx == INVALID_HANDLE_VALUE)
        {
            int err = Marshal.GetLastWin32Error();
            info.ErrorMessage = $"CreateActCtx 失败 (错误码 {err}): {new Win32Exception(err).Message}\n" +
                                $"清单路径: {manifestPath}";
            return info;
        }

        if (!ActivateActCtx(hActCtx, out IntPtr cookie))
        {
            int err = Marshal.GetLastWin32Error();
            ReleaseActCtx(hActCtx);
            info.ErrorMessage = $"ActivateActCtx 失败 (错误码 {err}): {new Win32Exception(err).Message}";
            return info;
        }

        try
        {
            var type = Type.GetTypeFromProgID(info.ProgId);
            if (type == null)
            {
                info.ErrorMessage = $"激活上下文中找不到 ProgID: {info.ProgId}\n" +
                                    "请检查组件清单中的 comClass progid 是否匹配";
                return info;
            }

            var result = CreatePluginFromType(info, type);

            if (result.IsLoaded)
            {
                // 成功时保存激活上下文句柄 (COM 对象生命期内保持有效)
                result.ActCtxHandle = hActCtx;
                result.ActCtxCookie = cookie;
                hActCtx = IntPtr.Zero; // 阻止 finally 释放
            }
            // 失败时 hActCtx 保持非零 → finally 会正确 Deactivate + Release
            // 避免泄漏的激活上下文干扰后续注册表回退策略
            return result;
        }
        finally
        {
            // 只有失败时才释放上下文
            if (hActCtx != IntPtr.Zero)
            {
                DeactivateActCtx(0, cookie);
                ReleaseActCtx(hActCtx);
            }
        }
    }

    /// <summary>
    /// 从 COM 类型创建插件实例并读取属性
    /// </summary>
    private static PluginInfo CreatePluginFromType(PluginInfo info, Type type)
    {
        try
        {
            var comObject = Activator.CreateInstance(type);
            if (comObject == null)
            {
                info.ErrorMessage = $"Activator.CreateInstance 返回 null: {info.ProgId}";
                return info;
            }

            dynamic plugin = comObject;

            info.ComObject = comObject;
            info.Name = (string)plugin.Name;
            info.Version = (string)plugin.Version;
            info.Description = (string)plugin.Description;
            info.IsLoaded = true;

            plugin.Initialize();
        }
        catch (Exception ex)
        {
            info.ErrorMessage = $"加载失败: {ex.Message}";
        }

        return info;
    }

    /// <summary>
    /// 执行插件
    /// </summary>
    public static string ExecutePlugin(PluginInfo plugin, string input)
    {
        if (!plugin.IsLoaded || plugin.ComObject == null)
            return "插件未加载";

        try
        {
            dynamic comObj = plugin.ComObject;
            return (string)comObj.Execute(input);
        }
        catch (Exception ex)
        {
            return $"执行错误: {ex.Message}";
        }
    }

    /// <summary>
    /// 卸载插件: 释放 COM 对象和激活上下文
    /// </summary>
    public static void UnloadPlugin(PluginInfo plugin)
    {
        if (plugin.ComObject != null)
        {
            try
            {
                dynamic comObj = plugin.ComObject;
                comObj.Shutdown();
            }
            catch
            {
                // 忽略 Shutdown 异常
            }

            Marshal.ReleaseComObject(plugin.ComObject);
            plugin.ComObject = null;
        }

        // 释放动态激活上下文
        if (plugin.ActCtxCookie != IntPtr.Zero)
        {
            DeactivateActCtx(0, plugin.ActCtxCookie);
            plugin.ActCtxCookie = IntPtr.Zero;
        }
        if (plugin.ActCtxHandle != IntPtr.Zero)
        {
            ReleaseActCtx(plugin.ActCtxHandle);
            plugin.ActCtxHandle = IntPtr.Zero;
        }

        plugin.IsLoaded = false;
    }

    public class PluginConfig
    {
        public List<PluginEntry> Plugins { get; set; } = new();
    }

    public class PluginEntry
    {
        public string ProgId { get; set; } = "";
        public string ManifestPath { get; set; } = "";
    }
}
