#!/bin/bash
# 用法：./build-manifest.sh https://<用户名>.github.io/<仓库名>
set -e
BASE="${1%/}"
[ -z "$BASE" ] && { echo "用法：$0 https://<用户名>.github.io/<仓库名>"; exit 1; }
DIR="$(cd "$(dirname "$0")" && pwd)"
PROVIDER="$(echo "$BASE" | sed -E 's#https://([^.]+)\..*#\1#')"
sed -e "s#__BASE__#$BASE#g" -e "s#__PROVIDER__#$PROVIDER#g" "$DIR/manifest.template.xml" > "$DIR/manifest.xml"
echo "已生成 manifest.xml（地址：$BASE）"
