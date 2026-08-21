# 설치와 환경

```bash
pipx install --force --python python3.12 ./skills/image-to-editable-ppt/cli
ln -sfn "$PWD/skills/image-to-editable-ppt" ~/.codex/skills/image-to-editable-ppt
editppt doctor --json
```

로컬 Codex, Python 3.10+ 및 macOS용 Microsoft PowerPoint가 필요합니다. PaddleOCR-VL은 선택 사항이며 `PADDLE_OCR_TOKEN` 또는 `~/.editppt/config.yaml`을 사용할 수 있습니다. OCR 저하는 명시되며 PowerPoint 실제 렌더가 실패하면 `ready`가 차단됩니다.
