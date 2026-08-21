# 예시 프롬프트

```text
$image-to-editable-ppt를 사용해 source.png를 객체 단위로 편집 가능한 PPTX로 재구성하세요. 원본 문구, 데이터, 그룹과 상대 위치를 유지하고 일반 콘텐츠는 네이티브 객체로 만드세요. 로고, 사진, 스크린샷과 복잡한 일러스트만 작은 이미지로 사용할 수 있습니다. 정확한 page.pptx를 Microsoft PowerPoint에서 실제 렌더링하고 확인한 뒤 result.json을 작성하세요.
```

여러 페이지는 페이지마다 독립적인 로컬 Codex 작업을 실행하고 성공한 `page.pptx`를 원본 순서대로 조립합니다.
