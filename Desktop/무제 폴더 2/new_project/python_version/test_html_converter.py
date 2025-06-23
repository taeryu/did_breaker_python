#!/usr/bin/env python3
"""
🧪 DSD Breaker HTML → Excel 변환기 테스트
DART 스타일 HTML 감사보고서 샘플 생성
"""

import os
from pathlib import Path

def create_sample_dart_html():
    """DART 스타일 감사보고서 HTML 샘플 생성"""
    
    html_content = '''
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>감사보고서 - 테스트용 DART HTML</title>
    <style>
        body { font-family: 'Malgun Gothic', Arial, sans-serif; margin: 20px; }
        h1, h2 { color: #2E5BBA; border-bottom: 2px solid #2E5BBA; padding-bottom: 5px; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
        th { background-color: #f4f4f4; font-weight: bold; text-align: center; }
        .number { text-align: right; }
        .center { text-align: center; }
        .indent1 { padding-left: 20px; }
        .indent2 { padding-left: 40px; }
    </style>
</head>
<body>
    <h1>재무상태표</h1>
    <p><strong>기업명:</strong> 테스트 주식회사</p>
    <p><strong>보고서 기준일:</strong> 2023년 12월 31일</p>
    
    <table>
        <thead>
            <tr>
                <th>계정과목</th>
                <th>제54기말(2023.12.31)</th>
                <th>제53기말(2022.12.31)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><strong>자산</strong></td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent1">유동자산</td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent2">현금및현금성자산</td>
                <td class="number">50,000,000</td>
                <td class="number">45,000,000</td>
            </tr>
            <tr>
                <td class="indent2">단기금융상품</td>
                <td class="number">30,000,000</td>
                <td class="number">25,000,000</td>
            </tr>
            <tr>
                <td class="indent2">매출채권</td>
                <td class="number">40,000,000</td>
                <td class="number">35,000,000</td>
            </tr>
            <tr>
                <td class="indent2">재고자산</td>
                <td class="number">25,000,000</td>
                <td class="number">20,000,000</td>
            </tr>
            <tr>
                <td class="indent1"><strong>유동자산 합계</strong></td>
                <td class="number"><strong>145,000,000</strong></td>
                <td class="number"><strong>125,000,000</strong></td>
            </tr>
            <tr>
                <td class="indent1">비유동자산</td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent2">투자자산</td>
                <td class="number">20,000,000</td>
                <td class="number">18,000,000</td>
            </tr>
            <tr>
                <td class="indent2">유형자산</td>
                <td class="number">80,000,000</td>
                <td class="number">75,000,000</td>
            </tr>
            <tr>
                <td class="indent2">무형자산</td>
                <td class="number">15,000,000</td>
                <td class="number">12,000,000</td>
            </tr>
            <tr>
                <td class="indent1"><strong>비유동자산 합계</strong></td>
                <td class="number"><strong>115,000,000</strong></td>
                <td class="number"><strong>105,000,000</strong></td>
            </tr>
            <tr>
                <td><strong>자산총계</strong></td>
                <td class="number"><strong>260,000,000</strong></td>
                <td class="number"><strong>230,000,000</strong></td>
            </tr>
            <tr>
                <td><strong>부채</strong></td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent1">유동부채</td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent2">매입채무</td>
                <td class="number">25,000,000</td>
                <td class="number">22,000,000</td>
            </tr>
            <tr>
                <td class="indent2">단기차입금</td>
                <td class="number">20,000,000</td>
                <td class="number">18,000,000</td>
            </tr>
            <tr>
                <td class="indent1"><strong>유동부채 합계</strong></td>
                <td class="number"><strong>45,000,000</strong></td>
                <td class="number"><strong>40,000,000</strong></td>
            </tr>
            <tr>
                <td class="indent1">비유동부채</td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent2">장기차입금</td>
                <td class="number">40,000,000</td>
                <td class="number">35,000,000</td>
            </tr>
            <tr>
                <td class="indent1"><strong>비유동부채 합계</strong></td>
                <td class="number"><strong>40,000,000</strong></td>
                <td class="number"><strong>35,000,000</strong></td>
            </tr>
            <tr>
                <td><strong>부채총계</strong></td>
                <td class="number"><strong>85,000,000</strong></td>
                <td class="number"><strong>75,000,000</strong></td>
            </tr>
            <tr>
                <td><strong>자본</strong></td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent1">자본금</td>
                <td class="number">100,000,000</td>
                <td class="number">100,000,000</td>
            </tr>
            <tr>
                <td class="indent1">이익잉여금</td>
                <td class="number">75,000,000</td>
                <td class="number">55,000,000</td>
            </tr>
            <tr>
                <td><strong>자본총계</strong></td>
                <td class="number"><strong>175,000,000</strong></td>
                <td class="number"><strong>155,000,000</strong></td>
            </tr>
            <tr>
                <td><strong>부채및자본총계</strong></td>
                <td class="number"><strong>260,000,000</strong></td>
                <td class="number"><strong>230,000,000</strong></td>
            </tr>
        </tbody>
    </table>

    <h2>포괄손익계산서</h2>
    
    <table>
        <thead>
            <tr>
                <th>계정과목</th>
                <th>제54기(2023.01.01~2023.12.31)</th>
                <th>제53기(2022.01.01~2022.12.31)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><strong>수익</strong></td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent1">매출</td>
                <td class="number">400,000,000</td>
                <td class="number">350,000,000</td>
            </tr>
            <tr>
                <td class="indent1">기타수익</td>
                <td class="number">5,000,000</td>
                <td class="number">3,000,000</td>
            </tr>
            <tr>
                <td><strong>수익 합계</strong></td>
                <td class="number"><strong>405,000,000</strong></td>
                <td class="number"><strong>353,000,000</strong></td>
            </tr>
            <tr>
                <td><strong>비용</strong></td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent1">매출원가</td>
                <td class="number">250,000,000</td>
                <td class="number">220,000,000</td>
            </tr>
            <tr>
                <td class="indent1">판매비와관리비</td>
                <td class="number">80,000,000</td>
                <td class="number">70,000,000</td>
            </tr>
            <tr>
                <td class="indent1">기타비용</td>
                <td class="number">10,000,000</td>
                <td class="number">8,000,000</td>
            </tr>
            <tr>
                <td><strong>비용 합계</strong></td>
                <td class="number"><strong>340,000,000</strong></td>
                <td class="number"><strong>298,000,000</strong></td>
            </tr>
            <tr>
                <td><strong>법인세비용차감전순이익</strong></td>
                <td class="number"><strong>65,000,000</strong></td>
                <td class="number"><strong>55,000,000</strong></td>
            </tr>
            <tr>
                <td>법인세비용</td>
                <td class="number">15,000,000</td>
                <td class="number">12,000,000</td>
            </tr>
            <tr>
                <td><strong>당기순이익</strong></td>
                <td class="number"><strong>50,000,000</strong></td>
                <td class="number"><strong>43,000,000</strong></td>
            </tr>
        </tbody>
    </table>

    <h2>현금흐름표</h2>
    
    <table>
        <thead>
            <tr>
                <th>계정과목</th>
                <th>제54기(2023.01.01~2023.12.31)</th>
                <th>제53기(2022.01.01~2022.12.31)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><strong>영업활동현금흐름</strong></td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent1">당기순이익</td>
                <td class="number">50,000,000</td>
                <td class="number">43,000,000</td>
            </tr>
            <tr>
                <td class="indent1">감가상각비</td>
                <td class="number">8,000,000</td>
                <td class="number">7,000,000</td>
            </tr>
            <tr>
                <td class="indent1">운전자본의 변동</td>
                <td class="number">(3,000,000)</td>
                <td class="number">(2,000,000)</td>
            </tr>
            <tr>
                <td><strong>영업활동순현금흐름</strong></td>
                <td class="number"><strong>55,000,000</strong></td>
                <td class="number"><strong>48,000,000</strong></td>
            </tr>
            <tr>
                <td><strong>투자활동현금흐름</strong></td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent1">유형자산 취득</td>
                <td class="number">(10,000,000)</td>
                <td class="number">(8,000,000)</td>
            </tr>
            <tr>
                <td class="indent1">투자자산 취득</td>
                <td class="number">(5,000,000)</td>
                <td class="number">(3,000,000)</td>
            </tr>
            <tr>
                <td><strong>투자활동순현금흐름</strong></td>
                <td class="number"><strong>(15,000,000)</strong></td>
                <td class="number"><strong>(11,000,000)</strong></td>
            </tr>
            <tr>
                <td><strong>재무활동현금흐름</strong></td>
                <td class="number"></td>
                <td class="number"></td>
            </tr>
            <tr>
                <td class="indent1">차입금의 증가</td>
                <td class="number">7,000,000</td>
                <td class="number">5,000,000</td>
            </tr>
            <tr>
                <td class="indent1">배당금 지급</td>
                <td class="number">(30,000,000)</td>
                <td class="number">(25,000,000)</td>
            </tr>
            <tr>
                <td><strong>재무활동순현금흐름</strong></td>
                <td class="number"><strong>(23,000,000)</strong></td>
                <td class="number"><strong>(20,000,000)</strong></td>
            </tr>
            <tr>
                <td><strong>현금및현금성자산의 순증가</strong></td>
                <td class="number"><strong>17,000,000</strong></td>
                <td class="number"><strong>17,000,000</strong></td>
            </tr>
            <tr>
                <td>기초 현금및현금성자산</td>
                <td class="number">33,000,000</td>
                <td class="number">16,000,000</td>
            </tr>
            <tr>
                <td><strong>기말 현금및현금성자산</strong></td>
                <td class="number"><strong>50,000,000</strong></td>
                <td class="number"><strong>33,000,000</strong></td>
            </tr>
        </tbody>
    </table>

    <h2>주석</h2>
    
    <h3>1. 회계정책</h3>
    <table>
        <thead>
            <tr>
                <th>항목</th>
                <th>회계정책</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>유형자산</td>
                <td>취득원가에서 감가상각누계액과 손상차손누계액을 차감한 가액으로 표시</td>
            </tr>
            <tr>
                <td>무형자산</td>
                <td>취득원가에서 상각누계액과 손상차손누계액을 차감한 가액으로 표시</td>
            </tr>
            <tr>
                <td>재고자산</td>
                <td>취득원가와 순실현가능가치 중 낮은 금액으로 측정</td>
            </tr>
        </tbody>
    </table>

    <h3>2. 중요한 회계추정 및 가정</h3>
    <table>
        <thead>
            <tr>
                <th>항목</th>
                <th>추정 및 가정</th>
                <th>금액(천원)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>유형자산 내용연수</td>
                <td>건물: 40년, 기계장치: 8년</td>
                <td>80,000,000</td>
            </tr>
            <tr>
                <td>무형자산 내용연수</td>
                <td>소프트웨어: 5년</td>
                <td>15,000,000</td>
            </tr>
            <tr>
                <td>충당부채</td>
                <td>과거 경험률 기준 추정</td>
                <td>2,000,000</td>
            </tr>
        </tbody>
    </table>

    <p><strong>작성일:</strong> 2024년 3월 15일</p>
    <p><strong>감사법인:</strong> 테스트 회계법인</p>
    
</body>
</html>
    '''
    
    return html_content

def create_simple_html():
    """간단한 테스트용 HTML"""
    
    html_content = '''
<!DOCTYPE html>
<html>
<head>
    <title>간단한 테이블 테스트</title>
    <meta charset="utf-8">
</head>
<body>
    <h1>매출 현황</h1>
    <table border="1">
        <tr>
            <th>월</th>
            <th>매출액</th>
            <th>전년 동기</th>
        </tr>
        <tr>
            <td>1월</td>
            <td>10,000,000</td>
            <td>9,500,000</td>
        </tr>
        <tr>
            <td>2월</td>
            <td>12,000,000</td>
            <td>11,000,000</td>
        </tr>
        <tr>
            <td>3월</td>
            <td>15,000,000</td>
            <td>13,500,000</td>
        </tr>
        <tr>
            <td><strong>1분기 합계</strong></td>
            <td><strong>37,000,000</strong></td>
            <td><strong>34,000,000</strong></td>
        </tr>
    </table>

    <h1>직원 현황</h1>
    <table border="1">
        <tr>
            <th>부서</th>
            <th>인원</th>
            <th>평균연봉</th>
        </tr>
        <tr>
            <td>영업팀</td>
            <td>15</td>
            <td>45,000,000</td>
        </tr>
        <tr>
            <td>개발팀</td>
            <td>20</td>
            <td>55,000,000</td>
        </tr>
        <tr>
            <td>관리팀</td>
            <td>8</td>
            <td>40,000,000</td>
        </tr>
    </table>
</body>
</html>
    '''
    
    return html_content

def test_converter_dependencies():
    """변환기에 필요한 라이브러리 테스트"""
    
    print("🧪 DSD Breaker HTML → Excel 변환기 의존성 테스트")
    print("="*60)
    
    dependencies = {
        'pandas': 'pandas',
        'beautifulsoup4': 'bs4',
        'xlsxwriter': 'xlsxwriter',
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
        print(f"\\n설치 명령어:")
        print(f"pip install {' '.join(missing)}")
        return False
    
    return True

def main():
    """메인 테스트 함수"""
    
    print("🔄 DSD Breaker HTML → Excel 변환기 테스트")
    print("="*60)
    
    # 1. 의존성 체크
    if not test_converter_dependencies():
        print("❌ 필수 라이브러리가 누락되었습니다.")
        return
    
    print("\\n✅ 모든 의존성 설치 완료")
    
    # 2. 테스트 HTML 파일 생성
    print("\\n📄 테스트 HTML 파일 생성 중...")
    
    # DART 스타일 감사보고서
    dart_html = create_sample_dart_html()
    with open('sample_dart_audit_report.html', 'w', encoding='utf-8') as f:
        f.write(dart_html)
    
    print("✅ sample_dart_audit_report.html 생성 완료")
    print("  🏢 DART 스타일 감사보고서")
    print("  📊 재무상태표, 손익계산서, 현금흐름표, 주석 포함")
    print("  🔢 실제 숫자 데이터와 한글 계정과목 포함")
    
    # 간단한 테스트 HTML
    simple_html = create_simple_html()
    with open('sample_simple_tables.html', 'w', encoding='utf-8') as f:
        f.write(simple_html)
    
    print("✅ sample_simple_tables.html 생성 완료")
    print("  📋 간단한 테이블 구조")
    print("  🧪 기본 변환 테스트용")
    
    # 3. 변환 테스트 안내
    print("\\n🎯 변환 테스트 포인트:")
    print("="*40)
    print("1. 🔄 HTML → Excel 변환 테스트:")
    print("   - DART 감사보고서: 복잡한 재무제표 구조")
    print("   - 단순 테이블: 기본 변환 기능")
    
    print("\\n2. 📊 변환 옵션 테스트:")
    print("   - 테이블별 시트 분리 ON/OFF")
    print("   - 데이터 정리 기능")
    print("   - 숫자 자동 인식")
    print("   - 원본 서식 유지")
    
    print("\\n3. 🎯 기대 결과:")
    print("   - 재무상태표, 손익계산서, 현금흐름표 각각 별도 시트")
    print("   - 숫자 데이터 자동 인식 (10,000,000 → 10000000)")
    print("   - 들여쓰기 및 계정과목 구조 보존")
    print("   - 한글 텍스트 완벽 처리")
    
    # 4. 실행 안내
    print("\\n🚀 DSD Breaker HTML → Excel 변환기 실행:")
    print("="*50)
    print("python3 dsd_breaker_converter.py")
    print()
    print("📁 테스트 파일:")
    print("  • sample_dart_audit_report.html (DART 감사보고서 스타일)")
    print("  • sample_simple_tables.html (간단한 테이블)")
    print()
    print("✅ HTML → Excel 변환 테스트 준비 완료!")
    print()
    print("💡 팁: 변환기에서 '미리보기' 기능으로 먼저 확인해보세요!")

if __name__ == "__main__":
    main()