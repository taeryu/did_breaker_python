#!/usr/bin/env python3
# 테스트용 스크립트 - 파일 경로 직접 지정

import sys
import os
sys.path.append('.')

# 메인 시스템 import
from Excel_템플릿_기반_결산보고서_생성_시스템 import MonthlyClosingProcessor

def test_system():
    """시스템 테스트 실행"""
    
    print("🧪 시스템 테스트 시작...")
    
    try:
        # 파일 경로들 설정
        mapping_file = "계정과목매핑표.xlsx"
        cell_mapping_file = "셀매핑.xlsx"
        sap_file = "시산표샘플.xlsx"
        template_files = ["경영실적요약템플릿.xlsx", "재무상태표템플릿.xlsx"]
        output_folder = "./test_reports"
        
        # 파일 존재 확인
        files_to_check = [mapping_file, cell_mapping_file, sap_file] + template_files
        for file_path in files_to_check:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
        
        print("✅ 모든 필요 파일 확인 완료")
        
        # 프로세서 생성 (파일 경로 직접 지정)
        print("🔧 프로세서 초기화 중...")
        processor = MonthlyClosingProcessor(mapping_file, cell_mapping_file)
        
        # 처리 실행
        print("🚀 월마감 처리 실행 중...")
        result = processor.process_monthly_closing(
            sap_file=sap_file,
            template_files=template_files,
            year=2024,
            month=11,
            output_folder=output_folder
        )
        
        # 결과 출력
        if result['success']:
            print("\n🎉 테스트 성공!")
            print(f"📁 출력 폴더: {output_folder}")
            print(f"📋 생성된 보고서: {len(result['created_reports'])}개")
            
            for report in result['created_reports']:
                print(f"   📄 {os.path.basename(report)}")
                
            # 요약 정보
            summary = result.get('summary', {})
            if summary:
                print("\n📊 재무 요약:")
                for key, value in summary.items():
                    print(f"   {key}: {value:,}원")
        else:
            print(f"\n❌ 테스트 실패: {result['error']}")
            
    except Exception as e:
        print(f"\n💥 예외 발생: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_system()