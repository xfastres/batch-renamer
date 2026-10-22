# 批量重命名工具

Windows 本地图形化批量改名工具。双击即可运行，不安装运行时、不联网、不改注册表。

当前版本：**1.1.0**

## 系统要求

- Windows 10 / 11
- Windows PowerShell 5.1（系统自带）
- 无需管理员权限

## 启动

1. 解压整个文件夹。
2. 双击 `启动.vbs`（推荐，不闪控制台）。
3. 备用：`启动.bat`，只启动同名的 `启动.vbs`。

`启动.vbs` 使用 `-ExecutionPolicy Bypass` 仅作用于本次进程，不能绕过组策略对脚本宿主的禁用。若被禁用，请用 PowerShell 直接运行主脚本。

## 核心功能

- **序号编号：** 序号命名、仅前缀、仅后缀、前缀+序号+原名；自然序/修改时间/创建时间/大小/随机排序。
- **过滤：** 文件/文件夹/两者；含子目录（只处理文件）；扩展名包含或排除；查找替换（作用于原名，不含扩展名）。
- **分段拼接：** 按连接符拆段编辑后重拼。
- **安全：** 两阶段改名；事务日志；循环/交换可撤销；部分失败可重试；不覆盖外部文件。
- **撤销：** 多步撤销，日志在 `%TEMP%\BatchRenamerLogs`。

## 数据安全

- 第一次文件系统变更前会写入事务清单。
- 启动时若发现未完成事务，会提示尝试恢复，**不会自动删除** `~rn_` 临时名。
- 旧版 `undo_*.csv` 仍可撤销。
- 自测使用独立临时目录，不改正式撤销记录。

## 命令行

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\BatchRenamer.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\BatchRenamer.ps1 -SelfTest
powershell -NoProfile -ExecutionPolicy Bypass -File .\BatchRenamer.ps1 -Diagnose -Folder D:\some-folder
```

自测通过退出码 0，失败非 0。诊断缺少 `-Folder` 时失败并返回非 0。

更多说明见 `操作说明.md`。变更见 `CHANGELOG.md`。

## 已知限制

- 递归模式只处理文件，不重命名子目录。
- 不支持仅大小写改名（Windows 默认大小写不敏感）。
- 轻量快照（路径/类型/大小）不能发现所有 TOCTOU 替换。
- 未完成多实例并发测试，不宣称可安全双开同一目录。
- DPI/小屏布局以人工检查为准。

## 测试

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Phase1Tests.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\BatchRenamer.ps1 -SelfTest
```
