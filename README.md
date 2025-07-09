# 📊 Excel MCP Server for Claude Desktop

클로드 데스크탑에서 Excel 파일을 처리할 수 있는 MCP (Model Context Protocol) 서버입니다.

## 🚀 기능

- ✅ **Excel 파일 읽기**: 다양한 형식의 Excel 파일 읽기 지원
- ✅ **Excel 파일 쓰기**: 데이터를 Excel 형식으로 저장
- ✅ **파일 정보 조회**: 시트 정보, 크기, 구조 분석
- ✅ **데이터 분석**: 통계 정보, 데이터 타입, 누락값 분석
- ✅ **데이터 필터링**: 조건에 따른 데이터 필터링

## 📦 설치

### 1. 의존성 설치
```bash
pip install -r requirements.txt
```

### 2. 클로드 데스크탑 설정

#### Windows:
1. `%APPDATA%\Claude\claude_desktop_config.json` 파일을 엽니다
2. 아래 설정을 추가합니다:

```json
{
  "mcpServers": {
    "excel-processor": {
      "command": "python",
      "args": ["C:\\path\\to\\your\\project\\excel_mcp_server.py"],
      "cwd": "C:\\path\\to\\your\\project",
      "env": {}
    }
  }
}
```

#### macOS:
1. `~/Library/Application Support/Claude/claude_desktop_config.json` 파일을 엽니다
2. 아래 설정을 추가합니다:

```json
{
  "mcpServers": {
    "excel-processor": {
      "command": "python3",
      "args": ["/path/to/your/project/excel_mcp_server.py"],
      "cwd": "/path/to/your/project",
      "env": {}
    }
  }
}
```

#### Linux:
1. `~/.config/claude/claude_desktop_config.json` 파일을 엽니다
2. macOS와 동일한 설정을 추가합니다.

### 3. 클로드 데스크탑 재시작

설정 파일을 수정한 후 클로드 데스크탑을 완전히 종료하고 다시 시작합니다.

## 🎯 사용 방법

클로드 데스크탑에서 다음과 같은 명령을 사용할 수 있습니다:

### Excel 파일 읽기
```
/path/to/sales_data.xlsx 파일을 읽어주세요
```

### 데이터 분석
```
/path/to/data.xlsx 파일의 통계 정보를 분석해주세요
```

### 데이터 필터링
```
/path/to/customers.xlsx에서 지역이 "서울"인 데이터만 필터링해주세요
```

### Excel 파일 생성
```
다음 데이터를 Excel 파일로 저장해주세요:
[{"이름": "김철수", "나이": 30}, {"이름": "이영희", "나이": 25}]
```

## 🛠️ 지원하는 도구

### 1. `read_excel`
- **설명**: Excel 파일을 읽어서 데이터를 반환합니다
- **매개변수**:
  - `file_path`: Excel 파일 경로 (필수)
  - `sheet_name`: 시트 이름 (선택)
  - `rows`: 읽을 행 수 제한 (선택)

### 2. `write_excel`
- **설명**: 데이터를 Excel 파일로 저장합니다
- **매개변수**:
  - `file_path`: 저장할 파일 경로 (필수)
  - `data`: 저장할 데이터 배열 (필수)
  - `sheet_name`: 시트 이름 (기본값: "Sheet1")

### 3. `get_excel_info`
- **설명**: Excel 파일의 기본 정보를 가져옵니다
- **매개변수**:
  - `file_path`: Excel 파일 경로 (필수)

### 4. `analyze_excel`
- **설명**: Excel 데이터를 분석하여 통계 정보를 제공합니다
- **매개변수**:
  - `file_path`: Excel 파일 경로 (필수)
  - `sheet_name`: 분석할 시트 이름 (선택)

### 5. `filter_excel_data`
- **설명**: Excel 데이터를 필터링합니다
- **매개변수**:
  - `file_path`: Excel 파일 경로 (필수)
  - `filters`: 필터 조건 객체 (필수)
  - `sheet_name`: 시트 이름 (선택)

## 🐛 문제 해결

### 1. 서버가 연결되지 않는 경우
- 파이썬 경로가 올바른지 확인
- `excel_mcp_server.py` 파일 경로가 정확한지 확인
- 의존성이 설치되었는지 확인: `pip list`

### 2. 로그 확인
- 프로젝트 폴더의 `excel_mcp.log` 파일을 확인

### 3. 권한 문제
- 스크립트 파일에 실행 권한이 있는지 확인 (Linux/macOS)

## 📝 예제

### 샘플 Excel 파일 생성 테스트
```python
python test_excel_mcp.py
```

이것은 테스트용 Excel 파일을 생성하고 MCP 서버 기능을 테스트합니다.

## 🔧 개발자 정보

이 MCP 서버는 다음 라이브러리를 사용합니다:
- `pandas`: 데이터 처리 및 분석
- `openpyxl`: Excel 파일 읽기/쓰기
- `xlsxwriter`: Excel 파일 생성 (선택적)

## 📄 라이선스

MIT License
