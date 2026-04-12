# 업무일지 자동화 툴

업무일지를 빠르게 작성하고 복사할 수 있는 브라우저 기반 도구입니다.

## 실행 방법

`index.html` 파일을 브라우저로 열면 바로 사용할 수 있습니다.

## 주요 기능

- 날짜와 작성자 입력
- 오늘 한 일 정리
- 이슈와 막힌 점 정리
- 내일 할 일 정리
- 생성된 업무일지 복사

## 다음에 추가하면 좋은 기능

- 작성 내용 자동 저장
- 일지 파일 다운로드
- 템플릿 여러 개 관리
- 팀 공유용 형식 추가

## 엑셀 자동화 실행

필요 패키지를 설치합니다.

```powershell
python -m pip install -r requirements.txt
```

기본 실행은 `low.xlsx`에서 가장 마지막 날짜를 찾아 `worklog_set.xlsx`의 4월 시트에 입력합니다.

```powershell
python automate_worklog.py
```

특정 날짜로 실행할 수도 있습니다.

```powershell
python automate_worklog.py --date 2026-04-11
```

결과 파일과 실행 리포트는 `output` 폴더에 저장됩니다.

## 웹앱 실행

브라우저에서 파일을 업로드하고 결과 엑셀을 바로 다운로드하려면 아래 명령을 실행합니다.

```powershell
python web_app.py
```

실행 후 브라우저에서 아래 주소로 접속합니다.

```text
http://127.0.0.1:5000
```

배포 서비스에서는 `Procfile`의 명령으로 실행할 수 있습니다.

## 로컬 업무용 실행

Python이 설치된 PC에서는 `start_app.bat`를 더블클릭하면 앱이 실행되고 브라우저가 열립니다.

기본 양식 파일은 앱 폴더의 `worklog_set.xlsx`를 자동으로 사용합니다. 평소에는 출결 파일 `low.xlsx`만 업로드하면 됩니다.

결과 파일은 기본적으로 `output` 폴더에 저장됩니다.

## EXE 만들기

Python이 없는 PC에서 쓰려면 `build_exe.bat`를 실행해 exe 파일을 만듭니다.

빌드가 끝나면 `dist` 폴더에 아래 파일을 같이 둡니다.

```text
dist/
└─ WorklogAutomation/
   ├─ WorklogAutomation.exe
   ├─ worklog_set.xlsx
   └─ _internal/
```

다른 PC에는 `dist/WorklogAutomation` 폴더를 통째로 전달하면 됩니다.
