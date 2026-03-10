using System.Runtime.InteropServices;

namespace ServiceCPlugin;

/// <summary>
/// ServiceC COM 插件接口 - 编码/哈希工具服务
/// 使用 InterfaceIsDual 同时支持 vtable 和 IDispatch
/// </summary>
[ComVisible(true)]
[Guid("B1A2C3D4-5E6F-7A8B-9C0D-E1F2A3B4C5D6")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IServiceCPlugin
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
