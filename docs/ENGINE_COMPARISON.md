# PowerPoint 없는 변환 엔진 검토

검토일: 2026-10-29. 현재 앱의 엔진은 PowerPoint COM이며, 아래 대안은 아직 앱에 통합하지 않았다.

## 요구사항

텍스트가 있는 도형을 이미지로 교체하면서 나머지 도형은 편집 가능해야 한다. 슬라이드 전체를 이미지나 PDF로 만드는 것만으로는 이 요구사항을 충족하지 않는다. 원본 글꼴을 사용할 수 있는 환경에서 회전, 겹침 순서, 그룹 및 슬라이드 밖 객체를 비교해야 한다.

## 후보 비교

| 엔진 | 확인된 기능 | 프로젝트 적용 시 확인할 사항 | 로컬 검증 상태 |
| --- | --- | --- | --- |
| PowerPoint COM | 현재 앱에서 도형 PNG 생성 및 PPT 저장 | PowerPoint 설치와 COM 작업 스레드 필요, 버전별 차이 | PowerPoint 2010에서 통합 테스트 통과 |
| Aspose.Slides for Python via .NET | Office 없이 파일 처리, 개별 도형 이미지 생성 | 이미지 좌표와 회전, 편집 객체 보존, 글꼴 처리, 배포 라이선스 | 라이브러리와 라이선스가 없어 실행 비교 미실시 |
| LibreOffice | headless 실행, 파일 형식 변환, 외부 API 제어 | 개별 도형 교체를 위한 별도 API 구현과 PPTX 왕복 저장 품질 검증 필요 | PATH 및 기본 설치 경로에서 실행 파일을 찾지 못해 비교 미실시 |

Aspose의 공식 문서는 개별 도형 이미지와 효과를 포함한 경계 계산을 제공한다고 설명한다. 다만 appearance 경계 이미지가 슬라이드 경계에 제한될 수 있으므로 슬라이드 밖 객체는 별도 검증이 필요하다. 평가판은 저장된 슬라이드에 워터마크를 추가하고 읽은 텍스트를 잘라내므로, 제한 없는 품질 비교에는 적절한 라이선스가 필요하다.

LibreOffice의 `--convert-to`는 파일 변환 기능이다. 이 명령만으로 텍스트 도형만 골라 교체할 수 있다고 가정하지 않는다. PowerPoint 없는 무료 대안을 선택하려면 개별 객체 제어의 구현 가능성을 먼저 확인한다.

## 비교 절차와 채택 조건

1. `tests/test_powerpoint.py`의 일반 텍스트, 37도 회전, 겹침, 중첩 그룹, 음수 좌표 사례를 기준 자료로 사용한다.
2. 동일한 글꼴과 동일한 출력 크기로 각 후보의 결과를 생성한다. 원본은 PowerPoint로 렌더링한 이미지를 기준으로 삼는다.
3. 전체 평균 픽셀 차이뿐 아니라 텍스트 줄바꿈, 콘텐츠 경계 위치, 겹침, 변환되지 않은 도형의 편집 가능성을 확인한다. 평균 차이가 작아도 국소적인 위치 오류는 채택 실패로 판단한다.
4. 표·차트·SmartArt를 변환하지 못할 때 누락을 명확히 보고하는지 확인한다. 실패 후 원본과 결과 파일의 상태도 검사한다.
5. 정확성이 확인된 뒤 변환 시간, 메모리, 배포 크기와 라이선스 비용을 비교한다. 측정 전에는 속도나 품질 우열을 단정하지 않는다.

현재 결론: 안정성 개선은 PowerPoint 엔진으로 검증했다. 설치 의존성 제거를 위한 다음 실험 후보는 도형 이미지 API가 있는 Aspose이며, LibreOffice는 개별 객체 교체 구현 가능성 확인이 선행되어야 한다. 어느 후보도 아직 실제 품질 비교를 통과한 것으로 간주하지 않는다.

## 공식 근거

- [PowerPoint Shape.Export](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.export)
- [PowerPoint ZOrderPosition](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.zorderposition)
- [Aspose의 Office 없는 파일 처리](https://docs.aspose.com/slides/features/without-powerpoint/)
- [Aspose 개별 도형 이미지](https://docs.aspose.com/slides/python-net/create-shape-thumbnails/)
- [Aspose 평가판과 라이선스 제한](https://docs.aspose.com/slides/python-net/licensing/)
- [LibreOffice 명령줄 및 headless 실행](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html)
