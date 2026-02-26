#!/usr/bin/env bash
set -euo pipefail

# ============================================
# Google Drive OAuth トークン更新スクリプト
# rclone authorize でトークンを取得し、
# GitHub Secrets (RCLONE_DRIVE_TOKEN) を更新する
# ============================================

REPO="keishimizu26629/rename_pdf"
SECRET_NAME="RCLONE_DRIVE_TOKEN"
GH_CONFIG_DIR="${HOME}/.config/gh-sub"

echo "=== Google Drive OAuth トークン更新 ==="
echo ""

# 前提チェック
for cmd in rclone gh; do
  if ! command -v "$cmd" &>/dev/null; then
    echo "エラー: ${cmd} がインストールされていません"
    echo "  brew install ${cmd}"
    exit 1
  fi
done

# gh認証チェック
if ! GH_CONFIG_DIR="$GH_CONFIG_DIR" gh auth status &>/dev/null; then
  echo "エラー: gh CLI が認証されていません"
  echo "  GH_CONFIG_DIR=$GH_CONFIG_DIR gh auth login"
  exit 1
fi

echo "ブラウザが開きます。Googleアカウントでログインしてください。"
echo ""

# rclone authorize でトークン取得
TOKEN=$(rclone authorize "drive" 2>&1 | grep -o '{.*}' | tail -1)

if [ -z "$TOKEN" ]; then
  echo "エラー: トークンの取得に失敗しました"
  exit 1
fi

echo ""
echo "トークン取得成功。GitHub Secrets を更新します..."

# GitHub Secrets に登録
GH_CONFIG_DIR="$GH_CONFIG_DIR" gh secret set "$SECRET_NAME" \
  --repo "$REPO" \
  --body "$TOKEN"

echo ""
echo "=== 完了 ==="
echo "  リポジトリ: ${REPO}"
echo "  Secret:     ${SECRET_NAME}"
echo "  更新日時:   $(date '+%Y-%m-%d %H:%M:%S')"
