MIMO 工资评分生成器 GitHub 自动打包完整包 V1.2

最省事用法：
1. 解压本 ZIP。
2. 新建一个干净 GitHub 仓库，或者清空旧仓库后重新上传。
3. 上传解压后的所有文件和文件夹，不要上传 ZIP 本身。
4. 进入 GitHub 仓库的 Actions 页面。
5. 等 build-desktop 跑完。
6. 下载 Artifacts：
   - MIMO-WageScore-Windows-EXE：Windows 给运营用。
   - MIMO-WageScore-macOS：Mac 用。

如果你继续使用已有仓库 mmgz：
- 这个包里同时提供了 .github/workflows/build.yml 和 .github/workflows/package_auto.yml。
- build.yml 是标准打包脚本。
- package_auto.yml 是防止旧 build.yml 无法覆盖时用的备用打包脚本。
- 如果 Actions 里旧 build.yml 还报红，但 package_auto.yml 成功，直接下载 package_auto.yml 那个任务生成的 Artifacts 即可。

注意：
- 不要只上传 ZIP 文件。
- 必须上传 server.py、wage_core.py、frontend、vendor、requirements-build.txt、.github/workflows 这些文件和文件夹。
- Windows 生成的是 exe 压缩包，下载后解压，双击 MIMO工资评分生成器.exe。
