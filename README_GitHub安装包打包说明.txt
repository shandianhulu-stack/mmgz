MIMO 工资评分生成器 V1.4 GitHub 自动打包说明

上传方式：
1. 解压本 ZIP。
2. 建议新建一个干净 GitHub 仓库，或清空旧仓库后上传。
3. 上传“解压后的所有文件和文件夹”，不要上传 ZIP 本身。
4. 仓库里只能保留一个工作流：.github/workflows/build-installer.yml。
5. 进入 Actions，等待 build-installer 跑完。

成功后下载：
- MIMO-WageScore-Windows-Installer-Setup
  里面是 MIMO_WageScore_Setup_V1.4.0.exe，Windows 安装包。

- MIMO-WageScore-Windows-Portable-Folder
  Windows 免安装备用版。

- MIMO-WageScore-macOS-Portable
  Mac 免安装备用版。

重要：
如果旧仓库里已有 .github/workflows/build.yml、package.yml、package_auto.yml，请先删除。
否则旧脚本会继续失败，看起来像新包失败。
