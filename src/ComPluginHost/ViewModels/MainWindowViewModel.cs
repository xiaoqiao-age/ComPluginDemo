using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ComPluginHost.Models;
using ComPluginHost.Services;

namespace ComPluginHost.ViewModels;

public partial class MainWindowViewModel : ObservableObject
{
    private static readonly string BaseDir = AppDomain.CurrentDomain.BaseDirectory;

    private static readonly string ConfigPath = Path.Combine(BaseDir, "plugins.json");

    [ObservableProperty]
    private ObservableCollection<PluginInfo> _plugins = new();

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ExecuteCommand))]
    [NotifyPropertyChangedFor(nameof(HasSelectedPlugin))]
    [NotifyPropertyChangedFor(nameof(PluginDetailText))]
    private PluginInfo? _selectedPlugin;

    [ObservableProperty]
    private string _inputText = string.Empty;

    [ObservableProperty]
    private string _outputText = string.Empty;

    [ObservableProperty]
    private string _statusText = "就绪 - 免注册 COM (SxS Manifest)";

    [ObservableProperty]
    private string _newProgId = string.Empty;

    [ObservableProperty]
    private string _newManifestPath = string.Empty;

    [ObservableProperty]
    private string _logText = string.Empty;

    public bool HasSelectedPlugin => SelectedPlugin?.IsLoaded == true;

    public string PluginDetailText
    {
        get
        {
            if (SelectedPlugin == null) return "请从左侧选择一个插件";
            if (!SelectedPlugin.IsLoaded) return $"插件加载失败:\n{SelectedPlugin.ErrorMessage}";
            return $"名称: {SelectedPlugin.Name}\n" +
                   $"版本: {SelectedPlugin.Version}\n" +
                   $"ProgID: {SelectedPlugin.ProgId}\n" +
                   $"清单: {(string.IsNullOrEmpty(SelectedPlugin.ManifestPath) ? "(无)" : SelectedPlugin.ManifestPath)}\n" +
                   $"加载方式: {SelectedPlugin.LoadSource}\n" +
                   $"描述: {SelectedPlugin.Description}";
        }
    }

    public MainWindowViewModel()
    {
        AppendLog("COM Plugin Host 启动 (Reg-Free COM / SxS Manifest)");
        AppendLog($"应用目录: {BaseDir}");
        LoadAllPlugins();
    }

    [RelayCommand]
    private void LoadAllPlugins()
    {
        foreach (var plugin in Plugins.Where(p => p.IsLoaded))
            ComPluginLoader.UnloadPlugin(plugin);
        Plugins.Clear();
        OutputText = string.Empty;

        var entries = ComPluginLoader.LoadPluginConfig(ConfigPath);
        AppendLog($"从 plugins.json 加载了 {entries.Count} 个插件配置");

        foreach (var entry in entries)
        {
            AppendLog($"加载: {entry.ProgId} (清单: {(string.IsNullOrEmpty(entry.ManifestPath) ? "exe.manifest" : entry.ManifestPath)})");
            var plugin = ComPluginLoader.LoadPlugin(entry, BaseDir);
            Plugins.Add(plugin);

            if (plugin.IsLoaded)
                AppendLog($"  [OK] {plugin.Name} v{plugin.Version} ({plugin.LoadSource})");
            else
                AppendLog($"  [FAIL] {plugin.ErrorMessage}");
        }

        // Pass Calculator COM object to VB6 StringProcessor
        PassCalculatorToVB6();

        // Pass ServiceC COM object to OrchestratorB
        PassServiceCToOrchestratorB();

        StatusText = $"已加载 {Plugins.Count(p => p.IsLoaded)}/{Plugins.Count} 个插件 (Reg-Free COM)";
    }

    /// <summary>
    /// 将 C# Calculator COM 对象传递给 VB6 StringProcessor 插件
    /// </summary>
    private void PassCalculatorToVB6()
    {
        var calculator = Plugins.FirstOrDefault(p =>
            p.IsLoaded && p.ProgId == "ComPluginDemo.Calculator");
        var vb6Plugin = Plugins.FirstOrDefault(p =>
            p.IsLoaded && p.ProgId == "VB6ComPlugin.StringProcessor");

        if (calculator == null || vb6Plugin == null)
            return;

        try
        {
            dynamic vb6Obj = vb6Plugin.ComObject!;
            vb6Obj.SetCalculatorObject(calculator.ComObject);
            AppendLog($"  [OK] 已将 Calculator COM 对象传递给 {vb6Plugin.Name}");

            // Pass base directory so VB6 can find CSharpComPlugin manifest
            // (needed for "Create via Manifest" button in VB6 IDE debug mode)
            vb6Obj.SetManifestBaseDir(BaseDir);
        }
        catch (Exception ex)
        {
            AppendLog($"  [WARN] 传递 Calculator 对象失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 将 ServiceC COM 对象传递给 OrchestratorB 插件
    /// </summary>
    private void PassServiceCToOrchestratorB()
    {
        var serviceC = Plugins.FirstOrDefault(p =>
            p.IsLoaded && p.ProgId == "ComPluginDemo.ServiceC");
        var orchestratorB = Plugins.FirstOrDefault(p =>
            p.IsLoaded && p.ProgId == "ComPluginDemo.OrchestratorB");

        if (serviceC == null || orchestratorB == null)
            return;

        try
        {
            dynamic orchestratorObj = orchestratorB.ComObject!;
            orchestratorObj.SetServiceCObject(serviceC.ComObject);
            AppendLog($"  [OK] 已将 ServiceC COM 对象传递给 {orchestratorB.Name}");
        }
        catch (Exception ex)
        {
            AppendLog($"  [WARN] 传递 ServiceC 对象失败: {ex.Message}");
        }
    }

    [RelayCommand]
    private void AddPlugin()
    {
        if (string.IsNullOrWhiteSpace(NewProgId)) return;

        var progId = NewProgId.Trim();
        if (Plugins.Any(p => p.ProgId == progId))
        {
            StatusText = $"插件 {progId} 已存在";
            return;
        }

        var entry = new ComPluginLoader.PluginEntry
        {
            ProgId = progId,
            ManifestPath = NewManifestPath.Trim()
        };

        AppendLog($"添加插件: {progId}");
        var plugin = ComPluginLoader.LoadPlugin(entry, BaseDir);
        Plugins.Add(plugin);

        // 保存配置
        var allEntries = Plugins.Select(p => new ComPluginLoader.PluginEntry
        {
            ProgId = p.ProgId,
            ManifestPath = p.ManifestPath
        }).ToList();
        ComPluginLoader.SavePluginConfig(ConfigPath, allEntries);

        if (plugin.IsLoaded)
        {
            AppendLog($"  [OK] {plugin.Name} v{plugin.Version}");
            StatusText = $"已添加: {plugin.Name}";
        }
        else
        {
            AppendLog($"  [FAIL] {plugin.ErrorMessage}");
            StatusText = $"加载失败: {progId}";
        }

        NewProgId = string.Empty;
        NewManifestPath = string.Empty;
    }

    [RelayCommand]
    private void RemovePlugin()
    {
        if (SelectedPlugin == null) return;

        var name = SelectedPlugin.DisplayName;
        ComPluginLoader.UnloadPlugin(SelectedPlugin);
        Plugins.Remove(SelectedPlugin);

        var allEntries = Plugins.Select(p => new ComPluginLoader.PluginEntry
        {
            ProgId = p.ProgId,
            ManifestPath = p.ManifestPath
        }).ToList();
        ComPluginLoader.SavePluginConfig(ConfigPath, allEntries);

        AppendLog($"已移除: {name}");
        StatusText = $"已移除: {name}";
        SelectedPlugin = null;
    }

    private bool CanExecute() => SelectedPlugin?.IsLoaded == true;

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private void Execute()
    {
        if (SelectedPlugin == null || !SelectedPlugin.IsLoaded) return;

        AppendLog($"[{SelectedPlugin.Name}] >>> {InputText}");
        var result = ComPluginLoader.ExecutePlugin(SelectedPlugin, InputText);
        OutputText = result;
        AppendLog($"[{SelectedPlugin.Name}] <<< {result}");
        StatusText = $"执行完成 - {SelectedPlugin.Name}";
    }

    private void AppendLog(string message)
    {
        var timestamp = DateTime.Now.ToString("HH:mm:ss");
        LogText += $"[{timestamp}] {message}\n";
    }
}
