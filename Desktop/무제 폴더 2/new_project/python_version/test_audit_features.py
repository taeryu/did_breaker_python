#!/usr/bin/env python3
"""
🧪 DSD Breaker 감사 기능 테스트
DART 감사보고서 검증 기능 테스트용 스크립트
"""

import pandas as pd
import numpy as np
import os
from pathlib import Path

def create_sample_financial_data():
    """감사보고서 스타일의 샘플 재무제표 생성"""
    
    # 재무상태표 샘플 데이터
    balance_sheet_data = {
        '항목': [
            '자산',
            '  유동자산',
            '    현금및현금성자산',
            '    단기금융상품', 
            '    매출채권',
            '    재고자산',
            '  유동자산 합계',
            '  비유동자산',
            '    투자자산',
            '    유형자산',
            '    무형자산',
            '  비유동자산 합계',
            '자산총계',
            '부채',
            '  유동부채',
            '    매입채무',
            '    단기차입금',
            '  유동부채 합계',
            '  비유동부채',
            '    장기차입금',
            '  비유동부채 합계',
            '부채총계',
            '자본',
            '  자본금',
            '  이익잉여금',
            '자본총계',
            '부채및자본총계'
        ],
        '2023': [
            '',  # 자산 (헤더)
            '',  # 유동자산 (헤더)
            50000,  # 현금및현금성자산
            30000,  # 단기금융상품
            40000,  # 매출채권
            25000,  # 재고자산
            145000,  # 유동자산 합계
            '',  # 비유동자산 (헤더)
            20000,  # 투자자산
            80000,  # 유형자산
            15000,  # 무형자산
            115000,  # 비유동자산 합계
            260000,  # 자산총계
            '',  # 부채 (헤더)
            '',  # 유동부채 (헤더)
            25000,  # 매입채무
            20000,  # 단기차입금
            45000,  # 유동부채 합계
            '',  # 비유동부채 (헤더)
            40000,  # 장기차입금
            40000,  # 비유동부채 합계
            85000,  # 부채총계
            '',  # 자본 (헤더)
            100000,  # 자본금
            75000,  # 이익잉여금
            175000,  # 자본총계
            260000   # 부채및자본총계
        ],
        '2022': [
            '',  # 자산 (헤더)
            '',  # 유동자산 (헤더)
            45000,  # 현금및현금성자산
            25000,  # 단기금융상품
            35000,  # 매출채권
            20000,  # 재고자산
            125000,  # 유동자산 합계
            '',  # 비유동자산 (헤더)
            18000,  # 투자자산
            75000,  # 유형자산
            12000,  # 무형자산
            105000,  # 비유동자산 합계
            230000,  # 자산총계
            '',  # 부채 (헤더)
            '',  # 유동부채 (헤더)
            22000,  # 매입채무
            18000,  # 단기차입금
            40000,  # 유동부채 합계
            '',  # 비유동부채 (헤더)
            35000,  # 장기차입금
            35000,  # 비유동부채 합계
            75000,  # 부채총계
            '',  # 자본 (헤더)
            100000,  # 자본금
            55000,  # 이익잉여금
            155000,  # 자본총계
            230000   # 부채및자본총계
        ]
    }
    
    # 손익계산서 샘플 데이터
    income_statement_data = {
        '항목': [
            '수익',
            '  매출',
            '  기타수익',
            '수익 합계',
            '비용',
            '  매출원가',
            '  판매비와관리비',
            '  기타비용',
            '비용 합계',
            '법인세비용차감전순이익',
            '법인세비용',
            '당기순이익'
        ],
        '2023': [
            '',  # 수익 (헤더)
            400000,  # 매출
            5000,   # 기타수익
            405000,  # 수익 합계
            '',  # 비용 (헤더)
            250000,  # 매출원가
            80000,   # 판매비와관리비
            10000,   # 기타비용
            340000,  # 비용 합계
            65000,   # 법인세비용차감전순이익
            15000,   # 법인세비용
            50000    # 당기순이익
        ],
        '2022': [
            '',  # 수익 (헤더)
            350000,  # 매출
            3000,   # 기타수익
            353000,  # 수익 합계
            '',  # 비용 (헤더)
            220000,  # 매출원가
            70000,   # 판매비와관리비
            8000,    # 기타비용
            298000,  # 비용 합계
            55000,   # 법인세비용차감전순이익
            12000,   # 법인세비용
            43000    # 당기순이익
        ]
    }
    
    # 의도적인 오류가 포함된 데이터셋 생성
    error_data = balance_sheet_data.copy()
    # 합계 오류 추가 (유동자산 합계를 의도적으로 틀리게)
    error_data['2023'][6] = 144000  # 실제: 145000
    
    # DataFrame 생성
    balance_sheet = pd.DataFrame(balance_sheet_data)
    income_statement = pd.DataFrame(income_statement_data)
    error_balance_sheet = pd.DataFrame(error_data)
    
    return balance_sheet, income_statement, error_balance_sheet

def create_sample_html():
    """감사보고서 스타일의 HTML 생성"""
    
    html_content = '''
    <!DOCTYPE html>
    <html>
    <head>
        <title>감사보고서 - 테스트용</title>
        <meta charset="utf-8">
    </head>
    <body>
        <h1>재무상태표</h1>
        <table border="1">
            <tr>
                <th>항목</th>
                <th>2023</th>
                <th>2022</th>
            </tr>
            <tr>
                <td><strong>자산</strong></td>
                <td></td>
                <td></td>
            </tr>
            <tr>
                <td>&nbsp;&nbsp;유동자산</td>
                <td></td>
                <td></td>
            </tr>
            <tr>
                <td>&nbsp;&nbsp;&nbsp;&nbsp;현금및현금성자산</td>
                <td>50,000</td>
                <td>45,000</td>
            </tr>
            <tr>
                <td>&nbsp;&nbsp;&nbsp;&nbsp;단기금융상품</td>
                <td>30,000</td>
                <td>25,000</td>
            </tr>
            <tr>
                <td>&nbsp;&nbsp;&nbsp;&nbsp;매출채권</td>
                <td>40,000</td>
                <td>35,000</td>
            </tr>
            <tr>
                <td>&nbsp;&nbsp;&nbsp;&nbsp;재고자산</td>
                <td>25,000</td>
                <td>20,000</td>
            </tr>
            <tr>
                <td>&nbsp;&nbsp;유동자산 합계</td>
                <td><strong>145,000</strong></td>
                <td><strong>125,000</strong></td>
            </tr>
            <tr>
                <td><strong>자산총계</strong></td>
                <td><strong>260,000</strong></td>
                <td><strong>230,000</strong></td>
            </tr>
        </table>
        
        <h1>손익계산서</h1>
        <table border="1">
            <tr>
                <th>항목</th>
                <th>2023</th>
                <th>2022</th>
            </tr>
            <tr>
                <td><strong>수익</strong></td>
                <td></td>
                <td></td>
            </tr>
            <tr>
                <td>&nbsp;&nbsp;매출</td>
                <td>400,000</td>
                <td>350,000</td>
            </tr>
            <tr>
                <td>&nbsp;&nbsp;기타수익</td>
                <td>5,000</td>
                <td>3,000</td>
            </tr>
            <tr>
                <td>&nbsp;&nbsp;수익 합계</td>
                <td><strong>405,000</strong></td>
                <td><strong>353,000</strong></td>
            </tr>
            <tr>
                <td><strong>당기순이익</strong></td>
                <td><strong>50,000</strong></td>
                <td><strong>43,000</strong></td>
            </tr>
        </table>
    </body>
    </html>
    '''
    
    return html_content

def test_audit_dependencies():
    """감사 기능에 필요한 라이브러리 테스트"""
    
    print("🧪 DSD Breaker 감사 기능 의존성 테스트")
    print("="*50)
    
    dependencies = {
        'pandas': 'pandas',
        'numpy': 'numpy',
        'beautifulsoup4': 'bs4',
        'openpyxl': 'openpyxl'
    }
    
    missing = []
    
    for package, import_name in dependencies.items():
        try:
            __import__(import_name)
            print(f"  ✅ {package}")
        except ImportError:
            print(f"  ❌ {package} - 설치 필요")
            missing.append(package)
    
    if missing:
        print(f"\n설치 명령어:")
        print(f"pip install {' '.join(missing)}")
        return False
    
    return True

def main():
    """메인 테스트 함수"""
    
    print("🔍 DSD Breaker 감사보고서 검증 기능 테스트")
    print("="*60)
    
    # 1. 의존성 체크
    if not test_audit_dependencies():
        print("❌ 필수 라이브러리가 누락되었습니다.")
        return
    
    print("\n✅ 모든 의존성 설치 완료")
    
    # 2. 샘플 재무제표 생성
    print("\n📊 샘플 재무제표 생성 중...")
    
    balance_sheet, income_statement, error_balance_sheet = create_sample_financial_data()
    
    # Excel 파일로 저장
    with pd.ExcelWriter('sample_financial_statements.xlsx') as writer:
        balance_sheet.to_excel(writer, sheet_name='재무상태표', index=False)
        income_statement.to_excel(writer, sheet_name='손익계산서', index=False)
        error_balance_sheet.to_excel(writer, sheet_name='오류포함재무상태표', index=False)
    
    print("✅ sample_financial_statements.xlsx 생성 완료")
    print(f"  📋 재무상태표: {len(balance_sheet)}개 항목")
    print(f"  📋 손익계산서: {len(income_statement)}개 항목")
    print(f"  📋 오류포함시트: 합계 불일치 오류 포함")
    
    # 3. 샘플 HTML 생성
    print("\n📄 샘플 HTML 보고서 생성 중...")
    
    html_content = create_sample_html()
    
    with open('sample_audit_report.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("✅ sample_audit_report.html 생성 완료")
    print("  🔍 재무상태표 및 손익계산서 테이블 포함")
    print("  📊 레벨 구조 (들여쓰기) 적용")
    
    # 4. 검증 포인트 안내
    print("\n🎯 테스트 검증 포인트:")
    print("="*40)
    print("1. 📊 레벨 감지 테스트:")
    print("   - Level 0: 자산, 부채, 자본 (대분류)")
    print("   - Level 1: 유동자산, 비유동자산 등 (중분류)")
    print("   - Level 2: 현금및현금성자산 등 (소분류)")
    
    print("\n2. ➕ 합계 검증 테스트:")
    print("   - 유동자산 합계: 50,000+30,000+40,000+25,000 = 145,000")
    print("   - 수익 합계: 400,000+5,000 = 405,000")
    print("   - 오류포함시트: 유동자산 합계 144,000 (오류)")
    
    print("\n3. 🔄 교차 참조 테스트:")
    print("   - 재무상태표와 손익계산서 간 공통 항목 확인")
    print("   - 당기순이익 일관성 검증")
    
    print("\n4. ⚠️ 오류 탐지 테스트:")
    print("   - 중복 항목 감지")
    print("   - 비정상적 음수/양수 패턴")
    print("   - 빈 셀 비율 분석")
    
    # 5. 실행 안내
    print("\n🚀 DSD Breaker 감사 도구 실행:")
    print("="*40)
    print("python3 dsd_breaker_audit.py")
    print()
    print("📁 테스트 파일:")
    print("  • sample_financial_statements.xlsx")
    print("  • sample_audit_report.html")
    print()
    print("✅ 테스트 준비 완료!")

if __name__ == "__main__":
    main()