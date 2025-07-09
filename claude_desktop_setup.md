
# 🖥️ 클로드 데스크탑 설정 가이드

## 1. 현재 프로젝트 경로
```
/Users/haeseoky/PycharmProjects/excel-mcp
```

## 2. 클로드 데스크탑 설정 파일 위치

### Windows:
```
%APPDATA%\Claude\claude_desktop_config.json
```

### macOS:
```
~/Library/Application Support/Claude/claude_desktop_config.json
```

### Linux:
```
~/.config/claude/claude_desktop_config.json
```

## 3. 설정 내용 (이 내용을 설정 파일에 추가하세요)

```json
{
  "mcpServers": {
    "excel-processor": {
      "command": "python3",
      "args": ["/Users/haeseoky/PycharmProjects/excel-mcp/excel_mcp_server.py"],
      "cwd": "/Users/haeseoky/PycharmProjects/excel-mcp",
      "env": {}
    }
  }
}
```

## 4. 설정 완료 후
1. 클로드 데스크탑을 완전히 종료
2. 클로드 데스크탑을 다시 시작
3. 새 대화에서 Excel 관련 작업 요청

## 5. 테스트 명령어 예시
```
현재 폴더의 sample_data.xlsx 파일을 읽어주세요
```
