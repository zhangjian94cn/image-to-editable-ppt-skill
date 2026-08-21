# Image to Editable PPT Skill

[简体中文](README.md) · [English](README_en.md) · **한국어**

슬라이드 이미지, 스캔 PDF, 이미지 기반 PPT/PPTX 페이지를 객체 단위로 편집 가능한 PowerPoint로 다시 만듭니다. 이 Fork는 `ningzimu/main@fb869763`을 기준으로 하며 페이지 하위 에이전트, Controller 상태 머신, 강제 승인 게이트 없이 하나의 Codex + Skill 흐름만 유지합니다.

## 사용법

Codex에 한 장의 슬라이드 이미지를 주고 Skill을 명시적으로 선택합니다.

```text
$image-to-editable-ppt를 사용하여 source.png를 객체 단위로 편집 가능한 한 장의 PowerPoint로 재구성하세요.
필요하면 OCR, 자산 추출, Builder, 미리보기 도구를 자율적으로 사용하세요.
source.png 전체를 슬라이드 배경이나 전체 덮개 이미지로 사용하지 마세요.
현재 페이지 디렉터리에 page.pptx와 result.json을 작성하세요.
```

하나의 Codex 작업이 한 페이지의 관찰, 제작, 필요 시 미리보기 수정까지 담당합니다.

필수 출력:

```text
source.png
page.pptx
result.json
preview.png   # Microsoft PowerPoint가 page.pptx를 실제 렌더링한 결과
```

## CLI

```bash
editppt prepare <input...>
editppt inspect text|layout|structure|pptx ...
editppt assets crop|separate|split-alpha|remove-chroma|brand ...
editppt build <page-dir>
editppt text-fit ...
editppt render <page-dir>
editppt compare <page-dir>
editppt assemble <page-dir...> --out <deck.pptx>
editppt formula render-latex ...
editppt doctor --json
```

CLI는 Codex가 필요할 때 선택하는 결정론적 보조 도구이며 강제 파이프라인이 아닙니다. 여러 페이지는 페이지별로 Codex 작업을 실행한 뒤 원본 순서대로 `editppt assemble`을 호출합니다.

공유 Builder는 설치된 글꼴을 해석하고 네이티브 연결선과 리치 텍스트
표를 작성하며 `editppt.authoring.SlideManifest`에서 재사용할 수 있습니다.
전체 슬라이드 래스터 이미지로 누락된 편집 구조를 숨기지 마세요. 로고,
사진, 지도, 스크린샷, 복잡한 일러스트는 독립적인 부분 이미지 객체로
사용할 수 있습니다. 정확한 `page.pptx`가 Microsoft PowerPoint에서 실제로
열리고 렌더링된 뒤에만 `result.json`을 작성할 수 있습니다.
