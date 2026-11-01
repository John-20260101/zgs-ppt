#!/usr/bin/env bash
# zgs-ppt 安装脚本：把技能链到各 CLI 的技能目录
#
#   ./install.sh                 # 软链到 ~/.agents|.codex|.claude/skills/zgs-ppt
#   ./install.sh ~/somewhere     # 指定技能安装路径（内部仍建 zgs-ppt 子目录）
#   ./install.sh --copy          # 复制而非软链（默认软链）
#   ./install.sh --uninstall     # 移除软链（不删源文件）
set -euo pipefail

SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
NAME="zgs-ppt"
MODE="link"
ROOT="$HOME/.agents/skills"
TARGET_DIRS=("$HOME/.agents/skills" "$HOME/.codex/skills" "$HOME/.claude/skills")

while [[ $# -gt 0 ]]; do
  case "$1" in
    --copy) MODE="copy"; shift ;;
    --uninstall) MODE="uninstall"; shift ;;
    -h|--help) sed -n '2,9p' "$0"; exit 0 ;;
    *) ROOT="$1"; shift ;;
  esac
done

DEST="$ROOT/$NAME"

if [[ "$MODE" == "uninstall" ]]; then
  for d in "${TARGET_DIRS[@]}"; do
    [[ -L "$d/$NAME" || -e "$d/$NAME" ]] && rm -rf "$d/$NAME" && echo "removed $d/$NAME"
  done
  exit 0
fi

if [[ "$ROOT" != "$HOME/.agents/skills" ]]; then
  mkdir -p "$ROOT"
  if [[ "$MODE" == "copy" ]]; then
    rm -rf "$DEST"; mkdir -p "$DEST"
    rsync -a --exclude '.git' --exclude '.DS_Store' --exclude '__pycache__' "$SRC/" "$DEST/"
    echo "copied → $DEST"
  else
    ln -sfn "$SRC" "$DEST"
    echo "linked → $DEST"
  fi
fi

for d in "${TARGET_DIRS[@]}"; do
  mkdir -p "$d"
  ln -sfn "$DEST" "$d/$NAME"
done

echo
echo "installed. 新开一个会话后生效（技能目录在会话启动时扫描）。"
echo "自检：python3 $DEST/scripts/build_pptx.py $DEST/examples/case-wang/王某某-outline.json --audit-only"
