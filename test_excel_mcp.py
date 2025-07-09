#!/usr/bin/env python3
"""
Excel MCP Server 테스트 스크립트
"""

import json
import asyncio
import subprocess
import tempfile
import pandas as pd
from pathlib import Path
import os

def create_sample_excel():
    """테스트용 샘플 Excel 파일 생성"""
    # 샘플 데이터 생성
    data = {
        '이름': ['김철수', '이영희', '박민수', '최지영', '정우진'],
        '나이': [28, 32, 24, 29, 35],
        '부서': ['개발팀', '마케팅팀', '개발팀', '인사팀', '마케팅팀'],
        '연봉': [4500, 3800, 4200, 3600, 5200],
        '입사일': ['2020-03-15', '2019-07-01', '2021-11-20', '2020-01-10', '2018-05-25']
    }
    
    df = pd.DataFrame(data)
    
    # 여러 시트가 있는 Excel 파일 생성
    sample_file = Path('sample_data.xlsx')
    with pd.ExcelWriter(sample_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='직원정보', index=False)
        
        # 부서별 통계 시트
        dept_stats = df.groupby('부서').agg({
            '이름': 'count',
            '나이': 'mean',
            '연봉': ['mean', 'sum']
        }).round(2)
        dept_stats.to_excel(writer, sheet_name='부서별통계')
    
    print(f"✅ 샘플 Excel 파일 생성됨: {sample_file}")
    return sample_file

async def test_mcp_server():
    """MCP 서버 기본 기능 테스트"""
    print("🧪 MCP 서버 테스트 시작...")
    
    # 샘플 파일 생성
    sample_file = create_sample_excel()
    
    print("\n📋 테스트 시나리오:")
    print("1. Excel 파일 정보 조회")
    print("2. Excel 데이터 읽기")
    print("3. Excel 데이터 분석")
    print("4. 데이터 필터링")
    print("5. 새로운 Excel 파일 생성")
    
    # 1. 파일 정보 테스트
    print("\n1️⃣ Excel 파일 정보 조회 테스트")
    info_message = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "get_excel_info",
            "arguments": {
                "file_path": str(sample_file)
            }
        }
    }
    print(f"요청: {json.dumps(info_message, ensure_ascii=False, indent=2)}")
    
    # 2. 데이터 읽기 테스트
    print("\n2️⃣ Excel 데이터 읽기 테스트")
    read_message = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {
            "name": "read_excel",
            "arguments": {
                "file_path": str(sample_file),
                "sheet_name": "직원정보"
            }
        }
    }
    print(f"요청: {json.dumps(read_message, ensure_ascii=False, indent=2)}")
    
    # 3. 데이터 분석 테스트
    print("\n3️⃣ Excel 데이터 분석 테스트")
    analyze_message = {
        "jsonrpc": "2.0",
        "id": 3,
        "method": "tools/call",
        "params": {
            "name": "analyze_excel",
            "arguments": {
                "file_path": str(sample_file),
                "sheet_name": "직원정보"
            }
        }
    }
    print(f"요청: {json.dumps(analyze_message, ensure_ascii=False, indent=2)}")
    
    # 4. 데이터 필터링 테스트
    print("\n4️⃣ 데이터 필터링 테스트")
    filter_message = {
        "jsonrpc": "2.0",
        "id": 4,
        "method": "tools/call",
        "params": {
            "name": "filter_excel_data",
            "arguments": {
                "file_path": str(sample_file),
                "sheet_name": "직원정보",
                "filters": {
                    "부서": "개발팀"
                }
            }
        }
    }
    print(f"요청: {json.dumps(filter_message, ensure_ascii=False, indent=2)}")
    
    # 5. 새 파일 생성 테스트
    print("\n5️⃣ 새로운 Excel 파일 생성 테스트")
    new_data = [
        {"제품명": "노트북", "가격": 1200000, "재고": 15},
        {"제품명": "마우스", "가격": 25000, "재고": 50},
        {"제품명": "키보드", "가격": 80000, "재고": 30}
    ]
    
    write_message = {
        "jsonrpc": "2.0",
        "id": 5,
        "method": "tools/call",
        "params": {
            "name": "write_excel",
            "arguments": {
                "file_path": "products.xlsx",
                "data": new_data,
                "sheet_name": "제품목록"
            }
        }
    }
    print(f"요청: {json.dumps(write_message, ensure_ascii=False, indent=2)}")
    
    print("\n✅ 테스트 시나리오 준비 완료!")
    print("\n📌 실제 MCP 서버 테스트 방법:")
    print("1. 의존성 설치: pip install -r requirements.txt")
    print("2. 서버 실행: python excel_mcp_server.py")
    print("3. 위의 JSON 메시지를 stdin으로 전달하여 테스트")
    
    return sample_file

def test_direct_functions():
    """MCP 서버 함수들을 직접 테스트"""
    print("\n🔬 직접 함수 테스트...")
    
    from excel_mcp_server import MCPServer
    
    async def run_direct_test():
        server = MCPServer()
        sample_file = create_sample_excel()
        
        print("\n📊 Excel 파일 정보:")
        info_result = await server.get_excel_info(str(sample_file))
        print(json.dumps(info_result, ensure_ascii=False, indent=2))
        
        print("\n📋 Excel 데이터 읽기:")
        read_result = await server.read_excel(str(sample_file), "직원정보")
        print(json.dumps(read_result, ensure_ascii=False, indent=2)[:500] + "...")
        
        print("\n📈 Excel 데이터 분석:")
        analyze_result = await server.analyze_excel(str(sample_file), "직원정보")
        print(json.dumps(analyze_result, ensure_ascii=False, indent=2)[:800] + "...")
        
        return sample_file
    
    return asyncio.run(run_direct_test())

def generate_claude_desktop_instructions():
    """클로드 데스크탑 설정 안내 생성"""
    current_path = os.path.abspath(".")
    python_path = "python" if os.name == 'nt' else "python3"
    
    instructions = f"""
# 🖥️ 클로드 데스크탑 설정 가이드

## 1. 현재 프로젝트 경로
```
{current_path}
```

## 2. 클로드 데스크탑 설정 파일 위치

### Windows:
```
%APPDATA%\\Claude\\claude_desktop_config.json
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
{{
  "mcpServers": {{
    "excel-processor": {{
      "command": "{python_path}",
      "args": ["{current_path}/excel_mcp_server.py"],
      "cwd": "{current_path}",
      "env": {{}}
    }}
  }}
}}
```

## 4. 설정 완료 후
1. 클로드 데스크탑을 완전히 종료
2. 클로드 데스크탑을 다시 시작
3. 새 대화에서 Excel 관련 작업 요청

## 5. 테스트 명령어 예시
```
현재 폴더의 sample_data.xlsx 파일을 읽어주세요
```
"""
    
    with open("claude_desktop_setup.md", "w", encoding="utf-8") as f:
        f.write(instructions)
    
    print("📄 클로드 데스크탑 설정 가이드가 생성되었습니다: claude_desktop_setup.md")
    print(instructions)

if __name__ == "__main__":
    print("🚀 Excel MCP Server 테스트 및 설정")
    print("=" * 50)
    
    # 직접 함수 테스트
    sample_file = test_direct_functions()
    
    # MCP 프로토콜 테스트 시나리오
    asyncio.run(test_mcp_server())
    
    # 클로드 데스크탑 설정 가이드 생성
    generate_claude_desktop_instructions()
    
    print(f"\n🎉 테스트 완료! 생성된 파일:")
    print(f"- {sample_file}")
    print(f"- claude_desktop_setup.md")
    print(f"- excel_mcp.log (실행 시 생성)")
