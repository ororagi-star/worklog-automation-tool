# 업무일지 자동 생성 도구

출결 엑셀 파일을 업로드하면 기본 업무일지 양식 2개에 기준일별 실적을 채워 결과 엑셀을 생성하는 Streamlit 앱입니다.

## 로컬 실행

```powershell
python -m pip install -r requirements.txt
python -m streamlit run streamlit_app.py
```

브라우저에서 아래 주소로 접속합니다.

```text
http://localhost:8501
```

## 웹 배포

Streamlit Community Cloud에 배포할 때는 다음 값으로 설정합니다.

```text
Repository: ororagi-star/worklog-automation-tool
Branch: main
Main file path: streamlit_app.py
```

앱은 저장소에 포함된 기본 양식 파일을 사용합니다.

```text
worklog_set1.xlsx
worklog_set2.xlsx
```

사용자가 업로드하는 `low.xlsx`는 저장소에 포함하지 않습니다.
