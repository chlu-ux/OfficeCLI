# Windows 构建与发布指南

本文说明如何在 Windows PowerShell 中测试 OfficeCLI，并生成 Linux ARM64、Alpine Linux ARM64 和 Windows 等平台的自包含发布二进制。

## 环境要求

- Windows 10 或 Windows 11
- PowerShell 5.1 或更高版本
- .NET 10 SDK

安装 SDK 后应重新打开 PowerShell，然后检查版本：

```powershell
dotnet --version
dotnet --info
```

本项目要求 .NET 10 SDK。示例验证版本为 `10.0.400`。

如果系统提示找不到 `dotnet`，可以在当前窗口临时补充安装目录：

```powershell
$env:Path = "C:\Program Files\dotnet;$env:Path"
dotnet --version
```

## 构建前测试

进入项目目录，还原依赖并运行 Release 测试：

```powershell
cd D:\data\OfficeCLI
dotnet restore .\officecli.slnx
dotnet test .\officecli.slnx -c Release --no-restore --nologo
```

只有测试全部通过后，才建议生成发布产物。

## PowerShell 脚本执行策略

如果执行 `build.ps1` 时出现“在此系统上禁止运行脚本”，建议只对当前 PowerShell 进程临时放行：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

该设置只影响当前窗口，关闭 PowerShell 后自动失效，不会永久降低系统执行策略。

也可以不修改当前会话，直接使用单次放行命令：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Runtime linux-arm64 -Configuration Release
```

不建议为了构建永久设置 `Unrestricted`。

## 生成 Linux ARM64 二进制

大多数使用 glibc 的 Linux 发行版，例如 Ubuntu、Debian、Rocky Linux 和 openEuler，使用：

```powershell
.\build.ps1 -Runtime linux-arm64 -Configuration Release
```

产物位置：

```text
bin\release\officecli-linux-arm64
```

这是自包含的单文件程序，目标 Linux 系统不需要另外安装 .NET Runtime。

## 生成 Alpine Linux ARM64 二进制

Alpine Linux 使用 musl libc，需要单独的目标：

```powershell
.\build.ps1 -Runtime linux-musl-arm64 -Configuration Release
```

产物位置：

```text
bin\release\officecli-linux-alpine-arm64
```

glibc 和 musl 产物不能混用。部署前应确认目标系统使用哪一种 libc。

## 生成其他平台二进制

例如生成 Windows x64 版本：

```powershell
.\build.ps1 -Runtime win-x64 -Configuration Release
```

生成脚本支持以下 RID：

| RID | 输出文件 |
| --- | --- |
| `linux-arm64` | `officecli-linux-arm64` |
| `linux-musl-arm64` | `officecli-linux-alpine-arm64` |
| `linux-x64` | `officecli-linux-x64` |
| `linux-musl-x64` | `officecli-linux-alpine-x64` |
| `win-x64` | `officecli-win-x64.exe` |
| `win-arm64` | `officecli-win-arm64.exe` |
| `osx-x64` | `officecli-mac-x64` |
| `osx-arm64` | `officecli-mac-arm64` |

一次生成全部平台产物：

```powershell
.\build.ps1 -All -Configuration Release
```

全平台构建耗时和磁盘占用较大。如果只需要部署到一个平台，推荐通过 `-Runtime` 只生成对应产物。

## 构建输出与校验

`build.ps1` 会输出以下信息：

- .NET SDK 版本
- 当前 Git commit
- 构建配置
- 目标 RID
- 产物路径
- SHA-256 校验值

也可以手动重新计算校验值：

```powershell
Get-FileHash .\bin\release\officecli-linux-arm64 -Algorithm SHA256
```

部署时应同时保存或传递 SHA-256，以便确认文件在复制过程中没有损坏。

## Linux ARM64 部署验证

把文件复制到目标 Linux ARM64 设备后，添加执行权限：

```bash
chmod +x officecli-linux-arm64
./officecli-linux-arm64 --version
```

执行基本 PPTX 冒烟测试：

```bash
./officecli-linux-arm64 create arm-smoke.pptx
./officecli-linux-arm64 validate arm-smoke.pptx --json
./officecli-linux-arm64 view arm-smoke.pptx outline --json
```

需要验证批处理时，可以继续执行项目所需的实际 batch 命令。发布验收不能只确认文件存在，必须在目标 ARM64 环境中真正启动二进制并操作 PPTX。

## 不使用构建脚本的发布方式

也可以直接调用 `dotnet publish`：

```powershell
dotnet publish .\src\officecli\officecli.csproj `
  -c Release `
  -r linux-arm64 `
  --self-contained true `
  -p:PublishAot=false `
  -o .\artifacts\linux-arm64
```

直接调用 `dotnet publish` 会保留完整发布目录。项目的 `build.ps1` 还负责 RID 白名单、统一命名、临时目录构建、原子替换以及 SHA-256 输出，因此正式生成项目约定的发布产物时优先使用 `build.ps1`。

## 常见问题

### 命令名称错误

正确命令是 `dotnet`，不是 `donet`。

### NuGet 依赖无法下载

确认当前机器能够访问：

```text
https://api.nuget.org/v3/index.json
```

如果公司网络使用代理或内部 NuGet 镜像，应按组织要求配置 `NuGet.Config`。

### Windows 不能运行 Linux 二进制

Windows 可以交叉编译 `linux-arm64`，但不能直接原生运行生成的 Linux ARM64 文件。最终冒烟测试应在真实 Linux ARM64 设备、ARM64 CI runner 或可靠的 ARM64 仿真环境中完成。
