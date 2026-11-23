# PPT / PIXEL — PPT 텍스트를 이미지로

PowerPoint의 텍스트가 포함된 도형을 PNG로 바꿔, 다른 PC에서 글꼴 차이로 발생하는 문제를 줄이는 Windows 데스크톱 도구입니다. 파일은 외부 서버에 업로드하지 않고 내 PC에서 처리합니다.

## 버전과 다운로드

**[v1.3 정식 출시](https://github.com/lahuman/ppt-text2img/releases/tag/1.3)** · Windows x64 · Microsoft PowerPoint 필요

[Windows 설치 프로그램](https://github.com/lahuman/ppt-text2img/releases/download/1.3/ppt-txt2img.exe) · [포터블 실행 파일](https://github.com/lahuman/ppt-text2img/releases/download/1.3/ppt-txt2img-portable.exe) · [SHA-256 체크섬](https://github.com/lahuman/ppt-text2img/releases/download/1.3/SHA256SUMS.txt)

- [소개 페이지](https://lahuman.github.io/ppt-text2img) / [페이지 소스](docs/index.html)
- [릴리즈 노트](RELEASE_NOTES.md)
- [릴리즈 절차](docs/RELEASING.md)

## v1.3 변경 사항

- 변환 전에 PPTX의 편집 제한 내장 글꼴 이름과 사용 슬라이드를 표시합니다. 대체를 선택하면 작업용 복사본에만 적용합니다. 대체 글꼴에 따라 줄바꿈과 배치가 달라질 수 있습니다. 구형 PPT는 사전 검사 대상에서 제외됩니다.
- 도형 처리와 파일 저장 실패를 구분해 로그에 기록합니다. 도형 오류에는 해당 슬라이드와 도형 번호를, PowerPoint 오류에는 제공된 HRESULT와 상세 메시지를 남깁니다.
- Windows 빌드에 Python 3.14.10, Pillow 13.0.0, PyInstaller 6.22.3 보안 업데이트를 반영했습니다.

## v1.2 변경 사항

- **새 작업 화면**: 네이비 사이드바, 밝은 작업 영역, 둥근 패널과 개선된 타이포그래피.
- **명확한 시작 흐름**: 파일 선택과 변환 시작을 분리하고, Ctrl+O·숫자 진행률·작은 창 스크롤 지원.
- **원본 보호**: 읽기 전용 복사본을 열고, 복제 도형에서 이미지를 만든 뒤 성공한 경우에만 원본 도형 교체.
- **배치 보정**: 회전을 이미지에 포함하고 겹침 순서를 복원하며, 슬라이드 밖 객체와 글자 오버행을 위한 내보내기 여백 반영.
- **부분 실패 안내**: 전체 완료와 일부 완료를 구분하고, 결과 옆의 JSON 보고서에 확인 필요 항목 기록.
- **실행 안정성**: 임시 도형 정리 실패 시 저장 중단, 변환 중 창 닫기 제한, 파일을 열 때 매크로 실행 차단. PowerPoint 경고는 파일을 열거나 저장하는 동안만 억제하고 도형을 변환하는 동안에는 기존 설정을 사용합니다.

### 실제 실행 화면

새 P 아이콘을 적용한 v1.2 Windows 포터블 실행 파일의 시작 화면입니다. 이미지를 선택하면 원본 크기로 볼 수 있습니다.

| 한국어 | English |
| --- | --- |
| [![한국어 시작 화면](docs/v1.2-ko.png)](docs/v1.2-ko.png) | [![English start screen](docs/v1.2-en.png)](docs/v1.2-en.png) |


## 실행 조건과 지원 범위

- Windows 10 / 11 및 설치된 Microsoft PowerPoint가 필요합니다.
- 변환할 PC에서 원본 글꼴을 사용할 수 있어야 합니다. 대상 PPT는 PowerPoint에서 닫아 주세요.
- 입력 형식은 `.ppt`, `.pptx`입니다. 한국어·영어 UI를 지원합니다.
- 텍스트가 있는 도형 전체가 이미지가 되어 텍스트 직접 편집·검색이 제한됩니다. 작업용 원본을 보관하고 전달용 사본으로 사용하세요.
- 중첩 그룹까지 검사해 변환할 텍스트가 있는 그룹만 해제합니다. 그림만 묶은 그룹은 결과 파일에서도 함께 이동하거나 편집할 수 있습니다. 표·차트·SmartArt는 자동 변환하지 않고 확인 필요 항목으로 보고합니다.
- 슬라이드 마스터·노트·애니메이션·하이퍼링크 보존은 지원 범위에 포함되지 않습니다. 결과 파일을 확인한 후 배포하세요.
- 이미지로 변환해도 글꼴의 사용·배포 라이선스 조건이 면제되는 것은 아닙니다.

## 사용 방법

1. **파일 찾아보기** 또는 **Ctrl+O**로 PPT를 선택합니다.
2. 선택한 경로를 확인하고 **이미지로 변환하기**를 누릅니다.
3. 확인 팝업의 주의 사항을 확인한 뒤 변환을 시작합니다.
4. 파일을 여는 동안 움직이는 진행 바가 표시됩니다. 변환이 시작되면 파일 선택 패널이 접히고 파일명, 현재 슬라이드와 도형 번호, 진행률이 표시됩니다. 창을 최소화해도 변환은 계속됩니다.
   PPTX에 보기·인쇄 전용 또는 편집이 제한된 내장 글꼴이 있으면 글꼴 이름과 직접 지정된 슬라이드 번호를 먼저 표시합니다. 공유 서식의 적용 가능 슬라이드는 따로 표시합니다. 설치된 대체 글꼴을 선택하면 임시 복사본에 적용해 변환하며, 취소하면 결과를 저장하지 않습니다. 글꼴 교체로 줄바꿈이 달라질 수 있습니다. 이 검사는 PowerPoint의 EOT 내장 글꼴을 대상으로 하며 구형 바이너리 PPT는 기존 방식으로 변환합니다.
5. 저장과 파일 확인이 끝나면 같은 화면에서 결과를 확인합니다.

| 결과 | 다음 작업 |
| --- | --- |
| 완료 | **결과 파일 열기**로 변환한 PPT를 엽니다. |
| 일부 완료 | **보고서 확인**으로 검토할 항목을 확인하거나 **결과 파일 열기**로 PPT를 엽니다. |
| 복구 필요 | **복구 파일 위치 열기**로 폴더를 열고, 화면에 표시된 파일을 확인한 뒤 다른 이름으로 저장합니다. |

보고서에는 `converted_count`와 `issues`가 기록됩니다. 슬라이드에서 발생한 문제에는 슬라이드 번호와 처리 단계가 포함되며, 도형 번호는 그룹 해제 후 처리 순서 기준입니다. 보고서 저장 자체가 실패하면 결과 화면에서 파일 경로와 문제 내용을 확인할 수 있습니다.

### 안정화 기능

- **취소 요청** 버튼으로 저장 시작 전까지 중단할 수 있습니다. 현재 실행 중인 PowerPoint COM 호출이 반환된 뒤 슬라이드·도형 경계에서 취소합니다. COM 자체가 멈추면 즉시 취소할 수 없으며, 사용자 PowerPoint를 강제 종료하지 않습니다.
- 같은 사용자 임시 폴더를 사용하는 앱 실행 간에는 한 번에 하나의 변환만 허용합니다. 다른 변환은 대기열에 넣지 않고 안내합니다. 사용자가 직접 PowerPoint를 조작하는 것까지 차단하지는 않습니다.
- 결과는 원본 폴더의 `.ppt-pixel-pending-*.ppt[x]`에 먼저 저장합니다. PowerPoint 닫기와 기본 파일 구조 검사를 마친 뒤 최종 이름으로 확정합니다. PPTX는 필수 ZIP 항목과 CRC, PPT는 OLE 헤더를 검사하며, 실제 PowerPoint 재개방이나 시각적 동일성을 보증하는 검사는 아닙니다.
- PowerPoint가 결과 문서를 닫지 못하면 결과 화면에 복구용 임시 파일 경로가 표시됩니다. 이 경우 보고서의 상태는 `recovery_required`입니다. 구조 검사나 최종 이름 전환 실패 시에는 오류창에 남은 파일 경로를 표시합니다.
- JSON 보고서는 임시 파일을 거쳐 교체합니다. 기존 `converted_count`, `issues`를 유지하며 `schema_version`, `job_id`, `state`, `metrics`를 추가합니다. 정리 단계 문제에는 슬라이드 번호가 없을 수 있습니다. PPT와 보고서 두 파일 전체가 하나의 트랜잭션으로 저장되는 것은 아닙니다.
- 렌더링 요청은 최대 2,500만 픽셀로 제한합니다. 이 값은 메모리 상한이나 성능 보장이 아니며, 문서별 실제 자원 사용은 별도 측정이 필요합니다.
- GUI 실행 시 `%LOCALAPPDATA%\PPT-Pixel\logs\app.log`에 순환 로그를 기록합니다. 파일당 1MiB, 백업 3개를 유지합니다. 도형·저장 오류는 작업 ID와 단계별로 기록합니다. 도형 오류에는 슬라이드·도형 번호가, COM 오류에는 제공된 HRESULT와 PowerPoint 상세 메시지가 포함됩니다. 로그와 JSON 보고서의 오류 메시지에 파일명이나 경로가 포함될 수 있으므로 공유 전에 확인하세요.

코드는 `app.py`(UI), `conversion.py`(변환·COM·저장), `translations.py`(번역), `runtime_support.py`(동시 실행 제한·로그)로 나뉩니다. 기존 `app.text_to_image_ppt()` 호출은 유지하며, 선택적 `cancel_event` 인자를 지원합니다.

### PNG 내보내기 성능

회전한 도형의 외곽과 실제 텍스트 범위를 기준으로 PNG 내보내기 영역을 줄이고 36pt 여백을 둡니다. 그림자·반사 등 효과가 있거나, 기존 영역의 경계가 소수점 좌표이거나, 필요한 속성을 읽을 수 없으면 기존 영역을 사용합니다.

PowerPoint 14.0에서 한영 텍스트 도형 24개가 있는 4장짜리 PPTX 한 파일을 개선 전후 각각 3회씩 교차 변환한 평균입니다.

| 측정 구간 | 개선 전 | 개선 후 |
| --- | ---: | ---: |
| PowerPoint PNG 내보내기 | 3.324초 | 1.391초 |
| PNG 자르기 및 다시 저장 | 0.500초 | 0.239초 |
| 슬라이드에 그림 삽입 | 0.141초 | 0.134초 |
| 전체 변환 | 6.137초 | 4.184초 |

이 문서에서 PNG 내보내기 시간은 58.1%, 전체 변환 시간은 31.8% 줄었습니다. 회전 텍스트 8개와 슬라이드 밖 도형 2개(한 도형은 두 조건에 해당)를 포함해 그림 24개 모두 개선 전후 PNG 크기·RGBA 픽셀·슬라이드 배치 좌표와 크기가 일치했습니다.

## 개발 및 테스트

실제 PowerPoint에서 텍스트 도형 60개를 한 슬라이드에 넣고 첫 도형 표시 후 창을 최소화했습니다. 이 실행에서 도형 진행 문구는 11.3초 동안 26회 갱신됐고 60개 모두 변환됐습니다. 창을 복원해 결과 파일 생성과 열기 버튼에 전달되는 경로를 확인했습니다. 실제 PowerPoint·Tk UI 검사를 포함해 로컬 테스트 40개가 통과했습니다.

보안 패치 빌드는 Python 3.14.10과 아래 의존성을 사용합니다. PowerPoint 2010은 호환성 테스트에 사용한 버전이며 보안 지원이 종료됐습니다. 실제 문서 변환에는 보안 업데이트를 받는 PowerPoint를 사용하세요.

```powershell
python -m venv venv
.\venv\Scripts\python.exe -m pip install --upgrade pip==26.3
.\venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\venv\Scripts\python.exe app.py
.\venv\Scripts\python.exe -m unittest discover -s tests -v
```

실제 PowerPoint 및 Tk UI 통합 테스트는 데스크톱 테스트 환경에서 선택적으로 실행합니다. 임시 PPT와 창을 생성하므로 사용자 문서를 작업 중인 PowerPoint가 없는 환경에서 실행하세요.

```powershell
$env:RUN_POWERPOINT_TESTS = "1"
$env:RUN_UI_TESTS = "1"
.\venv\Scripts\python.exe -m unittest discover -s tests -v
```

주요 의존성은 `pywin32`, `Pillow`, Python에 포함된 `tkinter`입니다. `requirements-dev.txt`에는 PyInstaller도 포함됩니다.

## 빌드와 릴리즈 준비

```powershell
# 설치 없는 실행 파일과 SHA256SUMS.txt 생성
.\scripts\build-release.ps1

# Inno Setup 컴파일러가 준비된 환경에서 설치 프로그램까지 생성
.\scripts\build-release.ps1 -Installer -InnoCompiler "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
```

결과는 `build/release-v1.3/`에 생성됩니다.

| 파일 | 용도 |
| --- | --- |
| `ppt-txt2img-portable.exe` | 설치 없이 실행하는 앱 · PowerPoint는 필요 |
| `ppt-txt2img.exe` | Inno Setup 설치 프로그램 · `-Installer` 사용 시 생성 |
| `SHA256SUMS.txt` | 해당 빌드 산출물의 SHA-256 체크섬 |

저장소의 `Output/ppt-txt2img.exe`는 이전 배포 파일입니다. 새 릴리즈에는 빌드 스크립트가 생성한 파일을 사용하세요. 실행 파일은 Git 커밋 대신 GitHub Release 첨부 파일로 배포합니다.

소개 페이지는 별도 빌드가 없는 정적 HTML/CSS/JavaScript입니다. 로컬 확인은 다음 명령을 사용합니다.

```powershell
.\venv\Scripts\python.exe -m http.server 8871 --bind 127.0.0.1 --directory docs
# 브라우저에서 http://127.0.0.1:8871/ 열기
```

게시 전후 버전·다운로드 링크 전환 순서는 [릴리즈 절차](docs/RELEASING.md)에 설명되어 있습니다. [PowerPoint 없는 엔진 검토](docs/ENGINE_COMPARISON.md)는 별도 실험 과제로 남아 있습니다.
