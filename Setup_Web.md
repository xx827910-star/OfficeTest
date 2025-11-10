# Open XML SDK DOCX 格式分析工具 - 完整部署指南

> 本指南基于实际生产环境中遇到的所有问题和解决方案编写，适用于 Claude Code Web 环境。

---

## 📋 目录

1. [环境说明](#环境说明)
2. [问题清单](#问题清单)
3. [完整部署流程](#完整部署流程)
4. [常见错误处理](#常见错误处理)
5. [代码适配指南](#代码适配指南)
6. [项目维护](#项目维护)

---

## 环境说明

### 测试环境
- **操作系统**: Linux 4.4.0 (Ubuntu-like)
- **Python**: 3.11.14
- **.NET SDK**: 8.0.415
- **Open XML SDK**: 3.1.0
- **工作目录**: `/home/user/OfficeTest`

### 已知限制
- NuGet 官方源可能无法直接访问（代理/网络问题）
- 需要手动处理依赖包下载
- Open XML SDK 3.x API 与 2.x 不兼容

---

## 问题清单

在实际部署中遇到的所有问题：

### ❌ 问题 1: NuGet 包恢复失败

**错误信息**:
```
error NU1301: Unable to load the service index for source https://api.nuget.org/v3/index.json.
The proxy tunnel request to proxy 'http://21.0.0.43:15004/' failed with status code '401'
```

**原因分析**:
- .NET 的 HTTP 客户端在某些环境中无法访问 NuGet 官方源
- 虽然 `curl` 可以访问，但 .NET 网络栈有不同行为
- 可能是代理设置、SSL 证书验证或 HTTP 处理器问题

**解决方案**: ✅ [见步骤 2.3](#23-手动下载-nuget-包)

---

### ❌ 问题 2: API 兼容性错误

**错误信息**:
```
error CS1061: 'OnOffValue' does not contain a definition for 'Val'
error CS1061: 'Table' does not contain a definition for 'TableProperties'
error CS1061: 'SectionProperties' does not contain a definition for 'PageSize'
```

**原因分析**:
- Open XML SDK 3.x 的 API 与 2.x 有重大变化
- 属性访问方式从直接属性改为 `GetFirstChild<T>()` 方法
- 许多在线示例代码基于 2.x 版本，不能直接使用

**解决方案**: ✅ [见代码适配指南](#代码适配指南)

---

### ❌ 问题 3: 缺少依赖包

**错误信息**:
```
error NU1101: Unable to find package DocumentFormat.OpenXml.Framework
error NU1101: Unable to find package System.IO.Packaging
```

**原因分析**:
- DocumentFormat.OpenXml 依赖多个包
- 需要递归下载所有依赖

**解决方案**: ✅ [见步骤 2.3](#23-手动下载-nuget-包)

---

## 完整部署流程

### 步骤 1: 安装 .NET SDK

#### 1.1 下载安装脚本

```bash
cd /home/user/OfficeTest
wget https://dot.net/v1/dotnet-install.sh -O dotnet-install.sh
chmod +x dotnet-install.sh
```

#### 1.2 安装 .NET 8.0

```bash
./dotnet-install.sh --channel 8.0
```

**预期输出**:
```
dotnet-install: Installed version is 8.0.415
dotnet-install: Installation finished successfully.
```

#### 1.3 配置环境变量

```bash
export PATH="$PATH:/root/.dotnet"
export DOTNET_ROOT=/root/.dotnet
```

#### 1.4 验证安装

```bash
/root/.dotnet/dotnet --version
```

**预期输出**: `8.0.415`

---

### 步骤 2: 创建项目

#### 2.1 创建控制台应用

```bash
cd /home/user/OfficeTest
/root/.dotnet/dotnet new console -n DocxFormatAnalyzer -f net8.0
cd DocxFormatAnalyzer
```

**预期输出**:
```
The template "Console App" was created successfully.
Restore succeeded.
```

#### 2.2 编辑项目文件

创建或修改 `DocxFormatAnalyzer.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="DocumentFormat.OpenXml" Version="3.1.0" />
  </ItemGroup>
</Project>
```

#### 2.3 手动下载 NuGet 包

**⚠️ 关键步骤** - 绕过网络问题

```bash
# 创建本地包目录
mkdir -p /tmp/nuget-packages
cd /tmp/nuget-packages

# 下载主包
curl -L -o DocumentFormat.OpenXml.3.1.0.nupkg \
  "https://www.nuget.org/api/v2/package/DocumentFormat.OpenXml/3.1.0"

# 下载依赖包 1
curl -L -o DocumentFormat.OpenXml.Framework.3.1.0.nupkg \
  "https://www.nuget.org/api/v2/package/DocumentFormat.OpenXml.Framework/3.1.0"

# 下载依赖包 2
curl -L -o System.IO.Packaging.8.0.0.nupkg \
  "https://www.nuget.org/api/v2/package/System.IO.Packaging/8.0.0"
```

**验证下载**:
```bash
ls -lh /tmp/nuget-packages/*.nupkg
```

**预期输出**:
```
-rw-r--r-- 1 root root  10M ... DocumentFormat.OpenXml.3.1.0.nupkg
-rw-r--r-- 1 root root 1.4M ... DocumentFormat.OpenXml.Framework.3.1.0.nupkg
-rw-r--r-- 1 root root 336K ... System.IO.Packaging.8.0.0.nupkg
```

#### 2.4 配置本地 NuGet 源

创建 `NuGet.config` 文件（在项目目录中）:

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="local" value="/tmp/nuget-packages" />
  </packageSources>
</configuration>
```

#### 2.5 恢复包

```bash
cd /home/user/OfficeTest/DocxFormatAnalyzer
export PATH="$PATH:/root/.dotnet"
export DOTNET_ROOT=/root/.dotnet
/root/.dotnet/dotnet restore
```

**预期输出**:
```
Determining projects to restore...
All projects are up-to-date for restore.
```

**❌ 如果失败**: 检查 NuGet.config 路径和包文件是否存在

---

### 步骤 3: 编写代码

#### 3.1 核心代码结构

你在分支中看到的`EnhancedProgram.cs`可能是我本地调试的版本,你需要修改成你所在的环境才能运行

---

### 步骤 4: 构建和运行

#### 4.1 构建项目

```bash
cd /home/user/OfficeTest/DocxFormatAnalyzer
export PATH="$PATH:/root/.dotnet"
export DOTNET_ROOT=/root/.dotnet
/root/.dotnet/dotnet build
```

**预期输出**:
```
Build succeeded.
    0 Warning(s)
    0 Error(s)
Time Elapsed 00:00:03.26
```

#### 4.2 运行程序

```bash
/root/.dotnet/dotnet run
```


## 常见错误处理

### 错误 1: 找不到 dotnet 命令

**症状**:
```bash
bash: dotnet: command not found
```

**解决**:
```bash
export PATH="$PATH:/root/.dotnet"
export DOTNET_ROOT=/root/.dotnet
```

---

### 错误 2: NuGet 恢复失败

**症状**:
```
error NU1301: Unable to load the service index
```

**解决**:
1. 确认 NuGet.config 存在且路径正确
2. 检查本地包目录：`ls /tmp/nuget-packages/*.nupkg`
3. 重新下载包（见步骤 2.3）
4. 确保 NuGet.config 只指向本地源

---

### 错误 3: 编译错误 - CS1061

**症状**:
```
error CS1061: 'Table' does not contain a definition for 'TableProperties'
```

**原因**: 使用了 Open XML SDK 2.x 的 API

**解决**: 使用 3.x API
```csharp
// ❌ 错误 (2.x API)
var tPr = table.TableProperties;

// ✅ 正确 (3.x API)
var tPr = table.GetFirstChild<TableProperties>();
```

---

### 错误 4: 文件不存在

**症状**:
```
❌ 错误: 文件不存在 - /home/user/OfficeTest/test.docx
```

**解决**:
1. 检查文件路径
2. 确认文件存在：`ls -l /home/user/OfficeTest/test.docx`
3. 修改 `Program.cs` 中的 `docxPath` 变量

---

## 代码适配指南

### Open XML SDK 2.x vs 3.x API 对照表

| 操作 | 2.x API (❌ 旧) | 3.x API (✅ 新) |
|------|----------------|----------------|
| 获取表格属性 | `table.TableProperties` | `table.GetFirstChild<TableProperties>()` |
| 获取页面大小 | `sectionPr.PageSize` | `sectionPr.GetFirstChild<PageSize>()` |
| 获取页边距 | `sectionPr.PageMargin` | `sectionPr.GetFirstChild<PageMargin>()` |
| 获取子元素 | 直接属性访问 | `element.GetFirstChild<T>()` |
| 获取所有子元素 | `element.Elements<T>()` | `element.Elements<T>()` (不变) |

### 关键 API 模式

#### ✅ 正确的 3.x 模式

```csharp
// 1. 获取单个子元素
var pageSize = sectionPr.GetFirstChild<PageSize>();
if (pageSize != null)
{
    var width = pageSize.Width?.Value;
}

// 2. 遍历所有子元素
foreach (var para in body.Elements<Paragraph>())
{
    // 处理段落
}

// 3. 查找后代元素
foreach (var section in body.Descendants<SectionProperties>())
{
    // 处理节
}

// 4. 空安全访问
var fontSize = runPr.FontSize?.Val?.Value ?? "Default";
```


---

## 项目维护

### Git 配置

创建 `.gitignore`:

```gitignore
dotnet-install.sh

# .NET build outputs
**/bin/
**/obj/
*.dll
*.pdb
*.cache
```

### 清理编译输出

```bash
# 从 git 移除已跟踪的编译文件
git rm -r --cached DocxFormatAnalyzer/bin DocxFormatAnalyzer/obj

# 清理本地编译文件
rm -rf DocxFormatAnalyzer/bin DocxFormatAnalyzer/obj
```


## 故障排查清单

遇到问题时按此顺序检查：

- [ ] 环境变量是否设置（PATH, DOTNET_ROOT）
- [ ] .NET SDK 是否安装成功（`dotnet --version`）
- [ ] NuGet 包是否下载（`ls /tmp/nuget-packages/`）
- [ ] NuGet.config 是否正确配置
- [ ] 项目文件语法是否正确（.csproj）
- [ ] 代码使用的是 3.x API（GetFirstChild）
- [ ] 测试文件是否存在（test.docx）


## 参考资源

- [Open XML SDK 官方文档](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk)
- [.NET 8.0 文档](https://docs.microsoft.com/en-us/dotnet/core/)
- [Office Open XML 标准](http://officeopenxml.com/)


