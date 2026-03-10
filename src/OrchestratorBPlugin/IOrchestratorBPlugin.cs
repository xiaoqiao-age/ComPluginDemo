using System.Runtime.InteropServices;

namespace OrchestratorBPlugin;

/// <summary>
/// OrchestratorB COM 插件接口 - 编排器，使用 ServiceC 进行编码/哈希操作
/// DispId 1-6 与 IComPlugin 一致，DispId 7 用于接收 ServiceC 对象
/// </summary>
[ComVisible(true)]
[Guid("D3C4E5F6-7A8B-9C0D-1E2F-A3B4C5D6E7F8")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IOrchestratorBPlugin
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

    /// <summary>
    /// 接收宿主传递的 ServiceC COM 对象
    /// </summary>
    [DispId(7)]
    void SetServiceCObject(object serviceCObj);
}
