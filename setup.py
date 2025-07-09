#!/usr/bin/env python3
"""
Excel MCP Server 설치 및 설정 스크립트
"""

import os
import sys
import json
import platform
from pathlib import Path
import subprocess

def install_dependencies():
    """필요한 패키지 설치"""
    print("📦 의존성 패키지 설치 중...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ 의존성 설치 완료")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 의존성 설치 실패: {e}")
        return False

def get_claude_config_path():
    """운영체제별 클로드 설정 파일 경로 반환"""
    system = platform.system()
    
    if system == "Windows":
        return Path(os.environ['APPDATA']) / "Claude" / "claude_desktop_config.json"
    elif system == "Darwin":  # macOS
        return Path.home() / "Library" / "Application Support" / "Claude" / "claude_desktop_config.json"
    else:  # Linux
        return Path.home() / ".config" / "claude" / "claude_desktop_config.json"

def create_claude_config():
    """클로드 데스크탑 설정 파일 생성/업데이트"""
    print("🔧 클로드 데스크탑 설정 중...")
    
    config_path = get_claude_config_path()
    current_path = Path.cwd().resolve()
    
    # Python 실행 파일 경로
    python_cmd = "python" if platform.system() == "Windows" else "python3"
    
    # 새로운 MCP 서버 설정
    new_server_config = {
        "excel-processor": {
            "command": python_cmd,
            "args": [str(current_path / "excel_mcp_server.py")],
            "cwd": str(current_path),
            "env": {}
        }
    }
    
    # 기존 설정 파일이 있는지 확인
    if config_path.exists():
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                existing_config = json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            existing_config = {}
    else:
        existing_config = {}
        # 설정 디렉토리 생성
        config_path.parent.mkdir(parents=True, exist_ok=True)
    
    # mcpServers 섹션이 없으면 생성
    if "mcpServers" not in existing_config:
        existing_config["mcpServers"] = {}
    
    # 새로운 서버 설정 추가
    existing_config["mcpServers"].update(new_server_config)
    
    # 설정 파일 저장
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(existing_config, f, indent=2, ensure_ascii=False)
        
        print(f"✅ 클로드 설정 완료: {config_path}")
        print(f"📍 프로젝트 경로: {current_path}")
        return True
        
    except Exception as e:
        print(f"❌ 설정 파일 생성 실패: {e}")
        return False

def test_server():
    """MCP 서버 기본 테스트"""
    print("🧪 MCP 서버 테스트 중...")
    try:
        # 테스트 스크립트 실행
        result = subprocess.run([sys.executable, "test_excel_mcp.py"], 
                              capture_output=True, text=True, encoding='utf-8')
        
        if result.returncode == 0:
            print("✅ 서버 테스트 성공")
            return True
        else:
            print(f"❌ 서버 테스트 실패: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ 테스트 실행 오류: {e}")
        return False

def main():
    """메인 설치 프로세스"""
    print("🚀 Excel MCP Server 설치 시작")
    print("=" * 50)
    
    # 1. 의존성 설치
    if not install_dependencies():
        print("❌ 설치 중단: 의존성 설치 실패")
        return False
    
    # 2. 클로드 설정
    if not create_claude_config():
        print("❌ 설치 중단: 클로드 설정 실패")
        return False
    
    # 3. 서버 테스트
    if not test_server():
        print("⚠️ 경고: 서버 테스트 실패 (설정은 완료됨)")
    
    print("\n🎉 설치 완료!")
    print("\n📋 다음 단계:")
    print("1. 클로드 데스크탑을 완전히 종료")
    print("2. 클로드 데스크탑을 다시 시작")
    print("3. 새 대화에서 Excel 관련 작업 테스트")
    print("\n💡 테스트 명령어 예시:")
    print("'현재 폴더의 sample_data.xlsx 파일을 읽어주세요'")
    
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
