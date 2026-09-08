# 考次链接大魔王 — macOS 自动打包（GitHub Actions）

仓库结构（相对位置不能变，`uom_kaoci_export.py` 必须在 `kaoci_dmw/` 的上级）：

```
kaoci_github_repo/
├── kaoci_dmw/
│   ├── main.py          ← 已含 macOS 适配（Chrome 路径/pgrep/open 目录/APP_DIR）
│   ├── mac_build.py     ← 打包脚本（也可本机 Mac 直接跑）
│   └── mac_build.sh
├── uom_kaoci_export.py  ← 核心模块
└── .github/workflows/mac-build.yml
```

## 使用步骤（全程约 5 分钟，首次）

### 1. 建 GitHub 私有仓库并推送

网页建库后，在本目录执行（把 `xxx` 换成你的用户名/仓库名）：

```bash
cd 本目录
git init
git add .
git commit -m "init"
git branch -M main
git remote add origin https://github.com/xxx/kaoci-mac-build.git
git push -u origin main
```

> 推送时会提示输入 GitHub 用户名 + Personal Access Token（Settings → Developer settings → Tokens → Generate new token，勾 repo）。

### 2. 触发打包

GitHub 仓库页面 → **Actions** 标签 → 左侧 **mac-build** → 右侧 **Run workflow** 绿色按钮 → 运行。

（push 到 main 也会自动触发。）

### 3. 下载 .app

构建成功后（约 3-6 分钟），Actions 运行页底部出现 **Artifacts** 区块：

- 点 **kaoci-mac-arm64** → 下载 zip
- 解压得到 `考次链接大魔王.app`

### 4. 启动（Gatekeeper 处理）

下载的 app 会被 macOS 标记隔离，首次双击可能打不开。任选：

```bash
xattr -d com.apple.quarantine "考次链接大魔王.app"
open "考次链接大魔王.app"
```

## 说明

- `macos-14` runner 是 **Apple Silicon (arm64)**，产物适配 M1~M4；Intel Mac 请把 workflow 里 `runs-on: macos-14` 改为 `macos-13`（x86_64）
- 登录态 `token.txt` 与调试 Chrome profile 在运行时自动落在 `~/Library/Application Support/考次链接大魔王/`（macOS 上 .app 内部只读，不能写旁文件）
- 首次使用仍要点一次「一键获取登录态」并在弹出的 Chrome 里登录 UOM
- 产物未做 Apple 公证，仅自用/内部使用；跨设备分发需右键→打开 或 `xattr` 解除隔离
