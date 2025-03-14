
# Flask Doc Generator

이 프로젝트는 **Flask**를 이용해 웹 브라우저에서 이미지를 업로드받고, **python-docx**로 Word 문서를 생성한 뒤 다운로드할 수 있도록 해주는 간단한 예시 애플리케이션입니다. 또한, **Bootstrap**으로 기본적인 UI를 꾸몄으며, PyInstaller를 통해 단일 EXE로 패키징할 수 있습니다.

---

## 주요 기능

1. **이미지 업로드**
    - 여러 이미지를 동시에 업로드할 수 있습니다.
    - 업로드된 이미지는 임시 폴더에 저장됩니다.

2. **Word 문서 생성**
    - `python-docx`를 사용하여, 업로드된 이미지를 표 형태로 배치하고, 캡션을 넣어 `output.docx`를 생성합니다.
    - 표의 행 높이를 고정하고, 이미지를 그 높이에 맞춰 축소하거나, 반대로 이미지를 기준으로 표 크기가 늘어나도록 설정할 수도 있습니다.

3. **문서 다운로드**
    - Flask의 `/download` 페이지에 접속하면 자동으로 `output.docx`가 다운로드됩니다.
    - 다운로드 후 “뒤로가기” 버튼을 통해 업로드 폼으로 다시 돌아갈 수 있습니다.

4. **UI (Bootstrap 적용)**
    - 업로드 폼, 결과 페이지, 다운로드 페이지에 **Bootstrap**을 적용하여 간단히 꾸몄습니다.
    - `.text-center`, `.btn`, `.container` 등의 클래스를 사용해 화면 중앙 정렬, 버튼 스타일 등을 설정했습니다.

5. **PyInstaller 패키징**
    - `pyinstaller --onefile main.py` 명령으로 단일 실행 파일(EXE)을 만들 수 있습니다.
    - EXE만 있으면 Python이나 pip를 설치하지 않아도 로컬 PC에서 애플리케이션을 실행할 수 있습니다.

---

## 프로젝트 구조 예시

```
my_flask_doc_app/
├── main.py            # Flask + python-docx + Bootstrap UI 코드
├── requirements.txt   # 의존성 목록 (Flask, python-docx, etc.)
└── README.md          # 사용 방법, 개요 등 설명
```

---

## 사전 준비

1. **Python 3.7+**
    - 가상환경(venv) 사용 권장
2. **의존성 설치**
   ```bash
   pip install -r requirements.txt
   ```
    - requirements.txt 예시:
      ```
      Flask==2.2.5
      python-docx==0.8.11
      ```

3. **(선택) PyInstaller**
    - 단일 EXE로 패키징하려면:
      ```bash
      pip install pyinstaller
      ```

---

## 실행 방법

1. **소스 실행 (개발용)**
   ```bash
   python main.py
   ```
    - 실행하면 Tkinter 창이 뜨고, “실행” 버튼을 누르면 브라우저가 자동으로 열립니다.
    - 업로드 폼에서 이미지를 선택한 뒤 “업로드” 버튼을 누르면, Word 문서가 생성되고 다운로드할 수 있습니다.

2. **단일 EXE 빌드 (배포용)**
   ```bash
   pyinstaller --onefile main.py
   ```
    - 빌드가 완료되면 `dist/` 폴더 안에 `main.exe`(Windows 기준)가 생성됩니다.
    - 해당 EXE 파일만 다른 PC에 전달해 실행하면, Python 없이도 구동됩니다.

---

## 사용 예시

1. **메인 페이지**
    - 브라우저 주소: http://127.0.0.1:5000
    - “이미지 파일 선택” → 파일 여러 장 선택 → “업로드 및 문서 생성” 클릭

2. **문서 생성 결과**
    - “문서 생성 완료! 총 N개의 이미지 업로드됨.” 메시지
    - “다운로드” 버튼을 클릭하면 `/download` 페이지로 이동 → 자동으로 `output.docx` 다운로드

3. **다운로드 페이지**
    - “파일 다운로드가 곧 시작됩니다...”
    - “뒤로가기” 버튼을 통해 업로드 폼으로 복귀 가능

---

## 주요 코드 설명

- **`main.py`**
    - `app = Flask(__name__)`로 Flask 앱 생성
    - `@app.route('/')` → 업로드 폼 페이지 (Bootstrap 적용)
    - `@app.route('/upload', methods=['POST'])` → 이미지 업로드 처리 + `python-docx`로 `output.docx` 생성
    - `@app.route('/download')` → 다운로드 안내 페이지, 자동으로 `/download_file` 이동
    - `@app.route('/download_file')` → `send_file("output.docx", as_attachment=True)`로 문서 전송
    - **Tkinter**: `main()`에서 간단한 윈도우를 띄워 “실행” 버튼 클릭 시 Flask 서버 & 브라우저 오픈

---

## 주의사항 / 한계

1. **이미지 크기**
    - 업로드된 이미지가 매우 크면 Word 문서 용량이 커지고, 생성 속도도 느려집니다.
    - 필요하다면 Pillow 등으로 사전 리사이즈를 적용할 수 있습니다.

2. **표 행 높이 고정**
    - WD_ROW_HEIGHT_RULE.EXACT가 지원되지 않는 구버전 `python-docx` 환경에서는 XML을 직접 수정해야 합니다.
    - 샘플 코드에서 `set_row_height()` 함수로 이 문제를 우회하고 있습니다.

3. **마지막 빈 페이지**
    - 모든 그룹(2장씩) 처리 후 `doc.add_page_break()`가 남아있으면 빈 페이지가 생길 수 있으므로, 마지막 그룹에선 page_break를 넣지 않도록 조건을 주거나, 마지막에 빈 문단을 제거해야 합니다.

4. **PyInstaller 동적 import 이슈**
    - 대부분의 Pure Python 라이브러리는 자동 패키징되지만, 특정 라이브러리는 `--hidden-import`가 필요할 수 있습니다.

---

## 라이선스

- 본 프로젝트 코드는 자유롭게 수정/배포 가능합니다.
- Python, Flask, python-docx, Bootstrap 등의 라이선스는 각각의 라이선스를 따릅니다.
