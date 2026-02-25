#!/bin/bash
MANIFEST_URL="https://vvihsesoftware.github.io/dc-expert/manifest.xml"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

echo "Установка DC Expert..."
mkdir -p "$WEF_DIR"
curl -s -o "$WEF_DIR/manifest.xml" "$MANIFEST_URL"

echo "Готово! Перезапустите Excel -> Вставка -> Мои надстройки (стрелочка вниз) -> DC Expert."