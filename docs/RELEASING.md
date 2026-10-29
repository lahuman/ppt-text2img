# v1.2 릴리즈 준비 및 게시 순서

현재 상태는 **출시 준비**이다. 빌드, 릴리즈 초안, 공개, 소개 페이지의 다운로드 전환 순서로 진행한다. 빌드 스크립트는 원격 저장소나 공개 릴리즈를 변경하지 않는다.

## 1. 최종 검증

프로젝트 루트에서 실행한다.

```powershell
$env:RUN_POWERPOINT_TESTS = "1"
$env:RUN_UI_TESTS = "1"
.\venv\Scripts\python.exe -m unittest discover -s tests -v
node --check docs/site.js
git diff --check
```

추가로 최신 PowerPoint에서 실제 배포할 문서의 글꼴, 회전, 겹침 및 일부 완료 경고를 확인한다. 자동 테스트는 PowerPoint 2010 기준이며, 모든 문서의 시각적 동일성을 보장하지 않는다.

## 2. 배포 파일 생성

`requirements-dev.txt`를 설치한 가상환경과 Inno Setup 컴파일러가 필요하다.

```powershell
.\scripts\build-release.ps1 -Installer -InnoCompiler "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
```

버전은 `app.py`의 `APP_VERSION = "1.2"`를 기준으로 읽는다. `inno-setup.iss`의 `MyAppVersion`과 `docs/site.js`의 `release.version`도 같은 값이어야 한다.

`build/release-v1.2/`에서 아래 파일을 확인한다.

- `ppt-txt2img.exe`: 설치 프로그램. 기존 배포와 같은 다운로드 파일명이다.
- `ppt-txt2img-portable.exe`: 설치 없는 실행 파일. PowerPoint 설치 의존성은 동일하다.
- `SHA256SUMS.txt`: 위 파일의 SHA-256 해시.

Inno Setup이 없으면 `-Installer` 없이 포터블 파일과 체크섬만 생성할 수 있다. **두 자산을 전제로 한 소개 페이지를 출시 상태로 바꾸려면 설치 프로그램도 먼저 준비해야 한다.** 빌드 스크립트는 업로드나 태그 생성을 수행하지 않는다.

## 3. Git과 릴리즈 초안

1. `git status`와 diff를 검토한다. 앱, 테스트, 의존성, 문서, 소개 페이지 및 빌드 스크립트가 포함되어야 한다. `build/`, `dist/` 및 가상환경은 커밋하지 않는다.
2. 변경 사항을 커밋하고 원격에 push한다. 기존 GitHub Pages 게시 설정을 유지한다. 이 시점의 페이지는 v1.2를 출시 준비로 표시하고 기존 v1.1 다운로드를 안내한다.
3. 해당 커밋을 가리키는 **`1.2` 태그**로 GitHub Release 초안을 만든다. 기존의 `1.0` 태그처럼 `v` 접두사가 없는 형식을 사용한다.
4. 위의 세 파일을 초안에 첨부하고, 본문에는 `RELEASE_NOTES.md`의 v1.2 부분을 사용한다. PowerPoint 설치 필요 및 지원 제한도 명시한다.
5. 초안에서 설치 프로그램·포터블 파일·체크섬의 이름과 버전을 확인하고 설치·실행을 검증한다.

## 4. 공개 및 소개 페이지 전환

릴리즈를 실제 공개한 후 다음 경로가 다운로드되는지 확인한다.

- `https://github.com/lahuman/ppt-text2img/releases/download/1.2/ppt-txt2img.exe`
- `https://github.com/lahuman/ppt-text2img/releases/download/1.2/ppt-txt2img-portable.exe`

이후 다음 내용을 함께 수정한다.

1. `docs/site.js`의 `release.published`를 `false`에서 `true`로 변경한다. 페이지의 출시 상태, 기본 다운로드, 포터블 링크와 구버전 안내가 함께 바뀐다. `stableVersion`도 `1.2`로 맞춘다.
2. README의 버전 표·다운로드·미출시 안내를 v1.2 출시 상태로 갱신한다.
3. 릴리즈 노트의 `출시 준비`와 미출시 문구를 실제 게시일·출시 상태로 바꾼다.
4. 위 문서와 페이지 변경을 push하고 GitHub Pages에서 한국어·영어 다운로드와 모바일 표시를 확인한다.

미출시 파일을 최신 다운로드처럼 안내하지 않도록 `published` 전환을 자산 공개보다 먼저 하지 않는다. 페이지의 UI 미리보기는 실제 스크린샷이 아닌 구성 예시로 계속 명시한다.
