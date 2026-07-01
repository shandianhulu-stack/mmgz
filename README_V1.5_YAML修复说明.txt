V1.5 修复说明

本版专门修复 GitHub Actions 报错：
Invalid workflow file: .github/workflows/build-installer.yml#L70

原因：V1.4 的 workflow 在 run: | 里面动态写 Inno Setup 脚本时，here-string 内容没有按 YAML 要求缩进，导致 YAML 语法错误。

V1.5 处理：
1. 不再在 workflow 里动态生成 Inno Setup 脚本。
2. 直接使用仓库内 installer/MIMO_WageScore_Setup.iss。
3. 只保留 .github/workflows/build-installer.yml 一个 workflow。

上传 GitHub 后，Actions 成功后下载：
MIMO-WageScore-Windows-Installer-Setup
里面是 Windows 下一步安装包：MIMO_WageScore_Setup_V1.4.0.exe
