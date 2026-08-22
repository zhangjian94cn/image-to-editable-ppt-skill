# 빠른 시작

Skill과 CLI를 설치하고 `editppt doctor --json`을 실행합니다. Codex에 한 장의 슬라이드 이미지를 제공하고 다음과 같이 요청합니다.

```text
$image-to-editable-ppt를 사용해 source.png를 객체 단위로 편집 가능한 한 장의 PowerPoint로 재구성하세요. 전체 원본 이미지를 배경이나 덮개로 사용하지 마세요. 정확한 page.pptx를 Microsoft PowerPoint에서 실제 렌더링하고 확인한 뒤 result.json을 작성하세요.
```

성공한 페이지에는 `source.png`, `page.pptx`, `preview.png`, `result.json`이 있습니다. 여러 페이지는 준비된 각 페이지마다 Codex 작업 하나를 실행하고 원본 순서대로 `editppt assemble`을 사용합니다.
