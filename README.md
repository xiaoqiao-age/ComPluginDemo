# ComPluginDemo

一个 **免注册 COM (Reg-Free COM)** 插件系统演示项目，使用 Windows SxS (Side-by-Side) 清单替代注册表，实现跨语言 COM 组件的加载和协作。无需 `regsvr32` / `regasm`，部署即用。

## 架构

```
ComPluginHost (A - 宿主, Avalonia GUI)
│
├── CSharpComPlugin    → Calculator        数学计算器
├── VB6ComPlugin       → StringProcessor   字符串处理器 (VB6)
├── ServiceCPlugin     → ServiceC          编码/哈希工具
└── OrchestratorBPlugin → OrchestratorB    编排器
│
├── 传递 Calculator → StringProcessor  (VB6 调用 C# 计算器)
└── 传递 ServiceC  → OrchestratorB     (B 组合调用 C 的功能)
```

## 技术栈

| 技术 | 用途 |
|------|------|
| C# .NET 8 | COM 宿主 + 3 个 C# COM 组件 |
| Avalonia 11.2.3 | 桌面 UI 框架 |
| VB6 | VB6 ActiveX DLL COM 组件 |
| Windows SxS Manifest | 免注册 COM 加载机制 |
| Win32 ActivationContext API | 动态 COM 清单加载 |

## 插件功能

| 插件 | ProgID | 命令示例 |
|------|--------|----------|
| Calculator | `ComPluginDemo.Calculator` | `3+4*(2-1)` → `7` |
| StringProcessor | `VB6ComPlugin.StringProcessor` | VB6 字符串处理 |
| ServiceC | `ComPluginDemo.ServiceC` | `base64enc Hello` → `SGVsbG8=` |
| OrchestratorB | `ComPluginDemo.OrchestratorB` | `encode-hash Hello` → Base64 + SHA256 |

**ServiceC 支持的命令：** `base64enc`, `base64dec`, `urlenc`, `urldec`, `sha256`, `md5`

**OrchestratorB 支持的命令：** `encode-hash`, `pipeline`, `direct`, `info`

## 构建

```bash
dotnet build ComPluginDemo.sln
```

> VB6ComPlugin 需在 VB6 IDE 中单独编译，仓库已包含预编译的 VB6ComPlugin.dll。

## 核心特点

- **免注册 COM** — SxS 清单声明代替注册表，xcopy 部署
- **跨语言互操作** — C# ↔ VB6 通过 COM IDispatch 晚绑定通信
- **SxS 依赖嵌套** — B 的 manifest 声明依赖 C，SxS 自动在子目录解析
- **双策略加载** — manifest 优先，注册表回退 (支持 VB6 IDE 调试)
- **全部 x86** — 兼容 VB6 32 位限制

## 输出目录结构

```
bin\Debug\net8.0-windows\
├── ComPluginHost.exe
├── CSharpComPlugin\              Calculator 插件
├── VB6ComPlugin\                 VB6 插件
│   └── CSharpComPlugin\          └── VB6 的 SxS 依赖
├── ServiceCPlugin\               ServiceC 插件
└── OrchestratorBPlugin\          OrchestratorB 插件
    └── ServiceCPlugin\            └── B 的 SxS 依赖
```

## 详细文档

完整的技术架构、COM 接口定义、GUID 清单、设计决策等详见 [详细介绍.md](详细介绍.md)。

开发过程记录详见 [日志.md](日志.md)。

## License

[MIT](LICENSE)
