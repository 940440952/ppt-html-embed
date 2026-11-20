#!/bin/bash
# 把 manifest.xml 装进 Mac 版 PowerPoint。换电脑时把 manifest.xml 和这个脚本一起拷过去运行即可。
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF"
cp "$DIR/manifest.xml" "$WEF/html-embed.manifest.xml"
echo "已安装到：$WEF"
echo "请完全退出 PowerPoint（Cmd+Q）后重新打开，然后：插入 → 加载项 → 我的加载项 → HTML 嵌入"
