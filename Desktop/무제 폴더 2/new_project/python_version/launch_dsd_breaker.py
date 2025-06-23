#!/usr/bin/env python3
"""
🚀 DSD Breaker Python Edition 런처
간편한 실행을 위한 래퍼 스크립트
"""

import sys
import os
from pathlib import Path

def check_dependencies():
    """필수 의존성 체크"""
    required_modules = {
        'pandas': 'pandas',
        'numpy': 'numpy', 
        'matplotlib': 'matplotlib',
        'openpyxl': 'openpyxl',
        'tkinter': 'tkinter (Python 표준 라이브러리)'
    }
    
    missing_modules = []
    
    for module, display_name in required_modules.items():
        try:
            if module == 'tkinter':
                import tkinter
            else:
                __import__(module)
        except ImportError:
            missing_modules.append(display_name)
    
    return missing_modules

def main():
    """메인 런처 함수"""
    print("🔧 DSD Breaker Python Edition")
    print("=" * 50)
    
    # 의존성 체크
    missing = check_dependencies()
    
    if missing:
        print("❌ 다음 모듈이 설치되지 않았습니다:")
        for module in missing:
            print(f"   - {module}")
        print("\n설치 방법:")
        print("   pip install -r requirements.txt")
        return 1
    
    print("✅ 모든 의존성이 설치되어 있습니다.")
    print()
    
    # DSD Breaker 실행
    try:
        from dsd_breaker_concept import DSDBreakerApp
        
        print("🚀 DSD Breaker 시작...")
        app = DSDBreakerApp()
        app.run()
        
    except ImportError as e:
        print(f"❌ DSD Breaker 모듈 로드 실패: {e}")
        return 1
    except Exception as e:
        print(f"❌ 예상치 못한 오류: {e}")
        return 1
    
    print("👋 DSD Breaker 종료")
    return 0

if __name__ == "__main__":
    sys.exit(main())