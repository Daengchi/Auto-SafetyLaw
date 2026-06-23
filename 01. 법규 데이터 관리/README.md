# 법규준수평가표 자동 생성기

국가법령정보센터 Open API를 활용하여 법령명만 입력하면 법규준수평가표 Excel 파일(`.xlsm`)을 자동 생성하는 CLI 도구입니다. 법률·시행령·시행규칙의 3단비교표를 기반으로 시트를 구성하고, VBA 매크로를 통해 해당 조문에 "O" 표시 시 4번 시트(법규준수평가)에 자동으로 행이 삽입됩니다.

---

## 사전 요구사항

- **Python 3.10 이상**
- **Microsoft Excel** (xlsm 파일 실행 및 VBA 매크로 사용)
- **국가법령정보센터 Open API 인증키**
  - 발급: [https://open.law.go.kr/LSO/main.do](https://open.law.go.kr/LSO/main.do)

---

## 설치 및 실행

### 1. 저장소 클론

```bash
git clone <저장소 URL>
cd ※\ 코딩딩
```

### 2. 패키지 설치

```bash
pip install -r requirements.txt
```

### 3. API 키 설정

프로젝트 루트에 `.env` 파일을 생성합니다.

```
LAW_API_KEY=여기에_인증키_입력
```

`.env.example` 파일을 복사해서 사용해도 됩니다.

```bash
copy .env.example .env
```

### 4. 실행

```bash
python main.py --name 산업안전보건법
python main.py --name 화학물질관리법 --output ./output
python main.py --name 환경영향평가법 --debug
```

| 옵션 | 설명 |
|------|------|
| `--name`, `-n` | 검색할 법령 이름 (필수) |
| `--output`, `-o` | 출력 디렉토리 (기본값: 현재 디렉토리) |
| `--debug` | 원본 XML을 `debug_*.xml`로 저장 (파싱 문제 진단용) |

---

## template.xlsm 최초 생성 (1회 필요)

VBA 매크로가 포함된 `src/template.xlsm` 파일이 없으면 xlsm 생성이 실패합니다.

### 방법 A — 자동 스크립트 (pywin32 필요)

```bash
pip install pywin32
python tools/make_template.py
```

> **사전 설정 필요**: Excel → 파일 → 옵션 → 보안 센터 → 보안 센터 설정 → 매크로 설정 → **"VBA 프로젝트 개체 모델에 대한 액세스 신뢰"** 체크

### 방법 B — 수동 (Excel VBE)

1. Excel에서 새 통합 문서 생성 → `.xlsm` 형식으로 `src/template.xlsm`에 저장
2. `Alt+F11` → VBE → 좌측 트리에서 `ThisWorkbook` 더블클릭
3. `tools/make_template.py` 안의 `VBA_CODE = """` ~ `"""` 사이 코드를 붙여넣기
4. 저장 (`Ctrl+S`)

---

## 출력 파일 구조

생성되는 파일: `법규준수평가표_{법령명}_{날짜}.xlsm`

| 시트 | 내용 |
|------|------|
| **1. 전체법령(3단)** | 법률·시행령·시행규칙 3단비교표 + 해당여부(D열) |
| **2. 독립 시행령** | 3단비교에 누락된 시행령 조문 + 해당여부(C열) |
| **3. 독립 시행규칙** | 3단비교에 누락된 시행규칙 조문 + 해당여부(C열) |
| **4. 법규준수평가** | VBA가 자동 관리하는 평가표 (No./법률/시행령/시행규칙/업무내용/주관부서/평가결과) |

### VBA 동작 방식

- 시트 1~3의 해당여부 셀에 **"O"** 입력(드롭다운 또는 복붙) → 4번 시트에 해당 행 자동 삽입
- **"O" 해제** (삭제 또는 다른 값 입력) → 4번 시트에서 해당 행 자동 삭제
- 삽입 순서는 원본 시트 순서 기준으로 정렬
- 4번 시트의 E~G열(업무내용·주관부서·평가결과)은 사용자가 직접 입력하며, 다른 행 추가·삭제 시에도 데이터가 밀리지 않음
- Excel에서 열 때 **"콘텐츠 사용(Enable Content)"** 버튼을 반드시 클릭해야 매크로가 활성화됨

---

## VS Code 권장 설정

### 추천 확장

- **Python** (ms-python.python) — 실행, 디버깅, 린팅
- **Pylance** — 타입 추론 및 자동완성

### `.env` 파일 설정

VS Code에서 디버그 실행 시 `.env` 파일을 자동으로 읽으려면 `.vscode/launch.json`을 생성합니다.

```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "name": "법규준수평가표 생성",
      "type": "debugpy",
      "request": "launch",
      "program": "${workspaceFolder}/main.py",
      "args": ["--name", "산업안전보건법", "--output", "./output"],
      "envFile": "${workspaceFolder}/.env",
      "console": "integratedTerminal"
    }
  ]
}
```

### 터미널에서 실행

VS Code 통합 터미널(`Ctrl+``)에서 바로 실행할 수 있습니다.

```bash
python main.py --name 산업안전보건법 --output ./output
```
