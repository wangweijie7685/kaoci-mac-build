#!/usr/bin/env bash
# macOS 一键打包：双击或在终端跑 ./mac_build.sh
# 等价于 python3 mac_build.py
set -e
cd "$(dirname "$0")"
python3 mac_build.py
echo
echo "下一步："
echo "  open 'dist/考次链接大魔王.app'"