# 표준 작업 흐름

1. 호출자가 `editppt prepare`로 입력을 순서가 있는 `source.png` 페이지로 분리합니다.
2. 각 페이지는 `$image-to-editable-ppt`를 사용하는 독립적인 로컬 Codex 작업 하나가 담당합니다.
3. Codex는 전체 페이지를 관찰한 뒤 필요한 OCR, 구조, 자산, Builder, 글꼴 도구를 선택합니다.
4. 텍스트, 숫자, 표, 타임라인, 컨테이너와 연결선은 네이티브 객체로 만들고, 로고·사진·지도·스크린샷·복잡한 일러스트만 작은 이미지로 유지합니다.
5. `page.pptx`를 만든 뒤 `editppt render`로 Microsoft PowerPoint 실제 결과를 보고 구체적인 문제를 수정합니다.
6. 최종 PPTX가 실제 렌더링된 뒤에만 `result.json`을 작성하며, 여러 페이지는 `editppt assemble`로 원본 순서대로 합칩니다.

page worker, Controller, session 복구, dispatch/record 상태, coverage/containment 또는 Hybrid 대체 경로는 없습니다.
