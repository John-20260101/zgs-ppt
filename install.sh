#!/usr/bin/env bash
# zgs-ppt 安装脚本：把技能链到各 CLI 的技能目录
#
#   ./install.sh                 # 软链到 ~/.agents|.codex|.claude/skills/zgs-ppt
#   ./install.sh ~/somewhere     # 指定技能安装路径（内部仍建 zgs-ppt 子目录）
#   ./install.sh --copy          # 复制而非软链（默认软链）
#   ./install.sh --uninstall     # 移除软链（不删源文件）
#
# 关键点：$DEST 必须先真实存在，三个客户端目录再指向 $DEST。
# 旧版在默认 ROOT 时跳过了建 $DEST 那一步，结果 ~/.agents/skills/zgs-ppt
# 变成指向自己的死链，三个目标目录全部读不到 SKILL.md。
set -euo pipefail

SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
NAME="zgs-ppt"
MODE="link"
ROOT="$HOME/.agents/skills"
EXPLICIT_ROOT=0
TARGET_DIRS=("$HOME/.agents/skills" "$HOME/.codex/skills" "$HOME/.claude/skills")

while [[ $# -gt 0 ]]; do
  case "$1" in
    --copy) MODE="copy"; shift ;;
    --uninstall) MODE="uninstall"; shift ;;
    -h|--help) sed -n '2,12p' "$0"; exit 0 ;;
    *) ROOT="$1"; EXPLICIT_ROOT=1; shift ;;
  esac
done

DEST="$ROOT/$NAME"

if [[ "$MODE" == "uninstall" ]]; then
  for d in "${TARGET_DIRS[@]}"; do
    [[ -L "$d/$NAME" || -e "$d/$NAME" ]] && rm -rf "$d/$NAME" && echo "removed $d/$NAME"
  done
  exit 0
fi

mkdir -p "$ROOT"

# $DEST 自己要先立住：默认路径软链回源仓库；自定义路径按 MODE 建。
if [[ "$EXPLICIT_ROOT" -eq 0 ]]; then
  ln -sfn "$SRC" "$DEST"
  echo "linked → $DEST -> $SRC"
elif [[ "$MODE" == "copy" ]]; then
  rm -rf "$DEST"; mkdir -p "$DEST"
  rsync -a --exclude '.git' --exclude '.DS_Store' --exclude '__pycache__' "$SRC/" "$DEST/"
  echo "copied → $DEST"
else
  ln -sfn "$SRC" "$DEST"
  echo "linked → $DEST -> $SRC"
fi

for d in "${TARGET_DIRS[@]}"; do
  [[ "$d" == "$ROOT" ]] && continue        # 默认路径时 $DEST 已经在位，别自指
  mkdir -p "$d"
  ln -sfn "$DEST" "$d/$NAME"
  echo "linked → $d/$NAME -> $DEST"
done

if [[ ! -r "$DEST/SKILL.md" ]]; then
  echo "✗ 安装失败：$DEST/SKILL.md 读不到（检查源目录或改用 --copy）" >&2
  exit 1
fi

echo
echo "installed. 新开一个会话后生效（技能目录在会话启动时扫描）。"
echo "自检：python3 $DEST/scripts/build_pptx.py $DEST/examples/case-wang/王某某-outline.json --audit-only --report /tmp/zgs-ppt-selftest.md"
