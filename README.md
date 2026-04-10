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
