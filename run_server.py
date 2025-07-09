#!/usr/bin/env python3
"""
Excel MCP Server 실행 스크립트
디버깅 및 개발용
"""

import subprocess
import sys
import os
from pathlib import Path

def check_dependencies():
    """의존성 확인"""
    try:
        import pandas
        import openpyxl
        print("✅ 모든 의존성이 설치되어 있습니다.")
        return True
    except ImportError as e:
        print(f"❌ 의존성 누락: {e}")
        print("다음 명령어로 설치하세요: pip install -r requirements.txt")
        return False

def run_server():
    """MCP 서버 실행"""
    if not check_dependencies():
        return False
    
    print("🚀 Excel MCP Server 시작...")
    print("📌 이 서버는 클로드 데스크탑과 stdin/stdout으로 통신합니다.")
    print("📌 직접 실행 시에는 JSON-RPC 메시지를 입력하세요.")
    print("📌 종료하려면 Ctrl+C를 누르세요.")
    print("-" * 50)
    
    try:
        # MCP 서버 실행
        subprocess.run([sys.executable, "excel_mcp_server.py"])
    except KeyboardInterrupt:
        print("\n👋 서버 종료")
    except Exception as e:
        print(f"❌ 서버 실행 오류: {e}")
        return False
    
    return True

if __name__ == "__main__":
    run_server()
