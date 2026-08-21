# Installation and environment

```bash
pipx install --force --python python3.12 ./skills/image-to-editable-ppt/cli
ln -sfn "$PWD/skills/image-to-editable-ppt" ~/.codex/skills/image-to-editable-ppt
editppt doctor --json
```

The page flow requires local Codex, Python 3.10+, and Microsoft PowerPoint for
macOS. PaddleOCR-VL is optional and can use `PADDLE_OCR_TOKEN` or the existing
`~/.editppt/config.yaml`. OCR degradation is explicit; PowerPoint absence or a
failed authoritative render blocks `ready`.

The development benchmark runner uses host access only with frozen trusted
inputs. The Miaobi product runtime does not import benchmark code.
