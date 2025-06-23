#!/usr/bin/env python3
"""
DSD Breaker Python 버전 테스트 스크립트
"""

import pandas as pd
import numpy as np
import os
from pathlib import Path

def create_sample_data():
    """테스트용 샘플 데이터 생성"""
    
    # 샘플 데이터 생성
    np.random.seed(42)
    
    data = {
        '날짜': pd.date_range('2024-01-01', periods=100, freq='D'),
        '매출액': np.random.normal(1000000, 200000, 100),
        '고객수': np.random.poisson(50, 100),
        '지역': np.random.choice(['서울', '부산', '대구', '인천', '광주'], 100),
        '제품군': np.random.choice(['A제품', 'B제품', 'C제품'], 100),
        '만족도': np.random.uniform(1, 5, 100),
        '온도': np.random.normal(20, 10, 100)
    }
    
    df = pd.DataFrame(data)
    
    # 샘플 Excel 파일 저장
    output_path = Path("sample_data.xlsx")
    df.to_excel(output_path, index=False)
    
    print(f"✅ 샘플 데이터 생성 완료: {output_path}")
    print(f"📊 데이터 크기: {df.shape[0]}행 x {df.shape[1]}열")
    print(f"📋 컬럼: {', '.join(df.columns.tolist())}")
    
    return output_path

def test_chart_dependencies():
    """차트 생성에 필요한 라이브러리 테스트"""
    
    print("🔍 차트 생성 라이브러리 테스트...")
    
    try:
        import matplotlib
        print(f"  ✅ matplotlib {matplotlib.__version__}")
    except ImportError:
        print("  ❌ matplotlib 설치 필요: pip install matplotlib")
        return False
    
    try:
        import seaborn
        print(f"  ✅ seaborn {seaborn.__version__}")
    except ImportError:
        print("  ⚠️ seaborn 선택사항: pip install seaborn")
    
    try:
        import numpy
        print(f"  ✅ numpy {numpy.__version__}")
    except ImportError:
        print("  ❌ numpy 설치 필요: pip install numpy")
        return False
    
    try:
        import pandas
        print(f"  ✅ pandas {pandas.__version__}")
    except ImportError:
        print("  ❌ pandas 설치 필요: pip install pandas")
        return False
    
    return True

def main():
    """메인 테스트 함수"""
    
    print("🧪 DSD Breaker Python 버전 테스트")
    print("="*50)
    
    # 1. 라이브러리 테스트
    if not test_chart_dependencies():
        print("❌ 필수 라이브러리가 누락되었습니다.")
        return
    
    print()
    
    # 2. 샘플 데이터 생성
    sample_file = create_sample_data()
    
    print()
    
    # 3. DSD Breaker 앱 실행 안내
    print("🚀 DSD Breaker 실행 방법:")
    print("  python dsd_breaker_concept.py")
    print()
    print("📊 테스트 데이터:")
    print(f"  파일: {sample_file}")
    print("  - 다양한 데이터 타입 포함")
    print("  - 차트 생성 테스트 가능")
    print("  - 한글 컬럼명 테스트")
    
    print()
    print("✅ 테스트 완료! DSD Breaker를 실행해보세요.")

if __name__ == "__main__":
    main()