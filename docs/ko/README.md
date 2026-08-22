# Image to Editable PPT Skill

슬라이드 이미지, 스캔 PDF 또는 시각 기반 PPTX 페이지를 객체 단위로 편집 가능한 PowerPoint로 재구성합니다.

한 페이지는 하나의 로컬 Codex 작업이 `$image-to-editable-ppt`와 Skill의 결정론적 도구를 사용해 담당합니다. `page.pptx`를 만든 뒤 정확한 파일을 Microsoft PowerPoint에서 실제 렌더링하고 확인한 후에만 `result.json`을 작성합니다. 여러 페이지는 호출자가 원본 순서대로 조립합니다.

## 문서

- [빠른 시작](/ko/quickstart.md)
- [설계와 경계](/ko/design.md)
- [설치](/ko/installation.md)
- [표준 작업 흐름](/ko/workflow.md)
- [자주 묻는 질문](/ko/faq.md)
- [예시 프롬프트](/ko/prompts.md)

Skill은 OCR과 레이아웃 증거, 원본 픽셀 자산, 설치 글꼴 맞춤, 리치 텍스트, 네이티브 표와 연결선, 재사용 가능한 authoring 컴포넌트, 대상 전용 PowerPoint 렌더와 관계 보존 조립을 제공합니다. page subagent, Controller 상태, containment 또는 전체 스크린샷 폴백은 없습니다.
