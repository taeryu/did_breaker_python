#!/usr/bin/env python3
# 간단한 테스트 스크립트

import pandas as pd
import os

def simple_test():
    """간단한 기능 테스트"""
    
    print("🧪 간단한 테스트 시작...")
    
    try:
        # 1. 계정매핑표 로드 테스트
        print("📊 계정매핑표 로드 테스트...")
        mapping_df = pd.read_excel("계정과목매핑표.xlsx")
        print(f"✅ 매핑표 로드 성공: {len(mapping_df)}건")
        print(f"   컬럼: {list(mapping_df.columns)}")
        
        # 2. 시산표 로드 테스트  
        print("\n📊 시산표 로드 테스트...")
        trial_df = pd.read_excel("시산표샘플.xlsx")
        print(f"✅ 시산표 로드 성공: {len(trial_df)}건")
        print(f"   컬럼: {list(trial_df.columns)}")
        
        # 3. 셀매핑 로드 테스트
        print("\n📊 셀매핑 로드 테스트...")
        cell_df = pd.read_excel("셀매핑.xlsx", sheet_name="셀매핑")
        print(f"✅ 셀매핑 로드 성공: {len(cell_df)}건")
        print(f"   컬럼: {list(cell_df.columns)}")
        
        # 4. 컬럼 매핑 CSV 테스트
        print("\n📊 컬럼매핑 CSV 테스트...")
        col_df = pd.read_csv("column_mapping.csv")
        print(f"✅ 컬럼매핑 로드 성공: {len(col_df)}건")
        col_mapping = dict(zip(col_df['원본컬럼명'], col_df['한글컬럼명']))
        print(f"   매핑: {col_mapping}")
        
        # 5. 데이터 처리 테스트
        print("\n🔄 데이터 처리 테스트...")
        
        # 컬럼명 변환
        trial_df.columns = trial_df.columns.astype(str).str.strip()
        for old_col, new_col in col_mapping.items():
            if old_col in trial_df.columns:
                trial_df = trial_df.rename(columns={old_col: new_col})
        
        print(f"   변환된 컬럼: {list(trial_df.columns)}")
        
        # 매핑 테스트
        if '계정코드' in trial_df.columns and '계정코드' in mapping_df.columns:
            trial_df['계정코드'] = trial_df['계정코드'].astype(str)
            mapping_df['계정코드'] = mapping_df['계정코드'].astype(str)
            
            merged_df = pd.merge(trial_df, mapping_df, on='계정코드', how='left')
            print(f"✅ 매핑 완료: {len(merged_df)}건")
            
            # 매핑되지 않은 계정 확인
            unmapped = merged_df[merged_df['보고서계정명'].isna()]
            if not unmapped.empty:
                print(f"⚠️ 매핑되지 않은 계정: {len(unmapped)}건")
            else:
                print("✅ 모든 계정 매핑 성공")
                
            # 재무제표별 집계 테스트
            if '잔액' in merged_df.columns and '재무제표구분' in merged_df.columns:
                merged_df['잔액'] = pd.to_numeric(merged_df['잔액'], errors='coerce').fillna(0)
                
                is_data = merged_df[merged_df['재무제표구분'] == 'IS']
                bs_data = merged_df[merged_df['재무제표구분'] == 'BS']
                
                print(f"   손익계산서 항목: {len(is_data)}건")
                print(f"   재무상태표 항목: {len(bs_data)}건")
                
                if not is_data.empty:
                    is_totals = is_data.groupby('보고서계정명')['잔액'].sum()
                    print(f"   IS 계정별 합계: {len(is_totals)}개")
                    for account, amount in is_totals.head(5).items():
                        print(f"     {account}: {amount:,}원")
                
        print(f"\n🎉 모든 테스트 완료!")
        
    except Exception as e:
        print(f"\n💥 테스트 실패: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    simple_test()