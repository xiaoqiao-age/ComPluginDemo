using System.Runtime.InteropServices;

namespace CSharpComPlugin;

/// <summary>
/// COM 插件接口 - 所有 COM 插件必须实现此接口
/// 使用 InterfaceIsDual 同时支持 vtable 和 IDispatch（VB6 兼容）
/// </summary>
[ComVisible(true)]
[Guid("8C5E2B1A-3F4D-4A6E-9B7C-1D2E3F4A5B6C")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IComPlugin
{
    [DispId(1)]
    string Name { get; }

    [DispId(2)]
    string Version { get; }

    [DispId(3)]
    string Description { get; }

    [DispId(4)]
    void Initialize();

    [DispId(5)]
    string Execute(string input);

    [DispId(6)]
    void Shutdown();
}
