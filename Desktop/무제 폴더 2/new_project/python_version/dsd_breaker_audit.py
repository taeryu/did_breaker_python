#!/usr/bin/env python3
"""
🔍 DSD Breaker Audit Edition
DART 감사보고서 자동 검증 도구 - Python 구현
"""

import pandas as pd
import numpy as np
import re
from pathlib import Path
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os

class DSDBreakAuditApp:
    """DART 감사보고서 검증용 DSD Breaker"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🔍 DSD Breaker - DART 감사보고서 검증 도구")
        self.root.geometry("1000x700")
        
        # 상태 변수
        self.current_file = None
        self.html_content = None
        self.extracted_tables = []
        self.verification_results = []
        
        self.setup_ui()
    
    def setup_ui(self):
        """사용자 인터페이스 설정"""
        
        # 메뉴바
        self.create_menubar()
        
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="🔍 DSD Breaker - DART 감사보고서 검증", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 파일 선택 영역
        file_frame = ttk.LabelFrame(main_frame, text="📁 DART 보고서 파일", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(file_frame, text="HTML 파일 열기", 
                  command=self.open_html_file).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(file_frame, text="Excel 파일 열기", 
                  command=self.open_excel_file).grid(row=0, column=1, padx=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="DART 보고서 파일을 선택해주세요")
        self.file_label.grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        
        # 검증 기능 영역
        verification_frame = ttk.LabelFrame(main_frame, text="🛠️ 감사보고서 검증 기능", padding="10")
        verification_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 검증 버튼들
        buttons = [
            ("🔍 테이블 추출", self.extract_tables),
            ("📊 레벨 자동 감지", self.detect_levels),
            ("➕ 합계 검증", self.verify_sums),
            ("🔄 교차 참조 확인", self.cross_reference_check),
            ("⚠️ 오류 탐지", self.detect_errors),
            ("📋 검증 리포트", self.generate_report)
        ]
        
        for i, (text, command) in enumerate(buttons):
            row = i // 3
            col = i % 3
            ttk.Button(verification_frame, text=text, command=command, 
                      width=18).grid(row=row, column=col, padx=5, pady=5)
        
        # 결과 표시 영역 (Notebook 위젯 사용)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 로그 탭
        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="📋 처리 로그")
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=80)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 테이블 탭
        table_frame = ttk.Frame(self.notebook)
        self.notebook.add(table_frame, text="📊 추출된 테이블")
        
        # 테이블 표시용 Treeview
        self.table_tree = ttk.Treeview(table_frame)
        table_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.table_tree.yview)
        self.table_tree.configure(yscrollcommand=table_scrollbar.set)
        
        self.table_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0), pady=5)
        table_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
        # 검증 결과 탭
        result_frame = ttk.Frame(self.notebook)
        self.notebook.add(result_frame, text="✅ 검증 결과")
        
        self.result_text = scrolledtext.ScrolledText(result_frame, height=15, width=80)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # 초기 메시지
        self.log_message("🎉 DSD Breaker DART 감사보고서 검증 도구를 시작합니다")
        self.log_message("📁 HTML 또는 Excel 파일을 열어서 검증을 시작하세요")
    
    def create_menubar(self):
        """메뉴바 생성"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 파일 메뉴
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="HTML 열기", command=self.open_html_file)
        file_menu.add_command(label="Excel 열기", command=self.open_excel_file)
        file_menu.add_separator()
        file_menu.add_command(label="결과 저장", command=self.save_results)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.root.quit)
        
        # 검증 메뉴
        verify_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="검증", menu=verify_menu)
        verify_menu.add_command(label="전체 검증 실행", command=self.run_full_verification)
        verify_menu.add_command(label="합계 검증만", command=self.verify_sums)
        verify_menu.add_command(label="레벨 감지만", command=self.detect_levels)
        
        # 도움말 메뉴
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도움말", menu=help_menu)
        help_menu.add_command(label="사용법", command=self.show_help)
        help_menu.add_command(label="정보", command=self.show_about)
    
    def open_html_file(self):
        """DART HTML 파일 열기"""
        file_path = filedialog.askopenfilename(
            title="DART HTML 파일 선택",
            filetypes=[("HTML files", "*.html *.htm"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.html_content = f.read()
                
                self.current_file = file_path
                file_name = os.path.basename(file_path)
                self.file_label.config(text=f"📄 {file_name}")
                
                self.log_message(f"✅ HTML 파일 로드 성공: {file_name}")
                self.log_message(f"📊 HTML 내용 크기: {len(self.html_content):,} 문자")
                
                # 자동으로 테이블 추출 시작
                self.extract_tables()
                
            except Exception as e:
                messagebox.showerror("오류", f"HTML 파일 읽기 실패: {str(e)}")
                self.log_message(f"❌ HTML 파일 읽기 실패: {str(e)}")
    
    def open_excel_file(self):
        """Excel 파일 열기"""
        file_path = filedialog.askopenfilename(
            title="Excel 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # Excel 파일을 pandas로 읽기
                excel_data = pd.read_excel(file_path, sheet_name=None)  # 모든 시트 읽기
                
                self.current_file = file_path
                file_name = os.path.basename(file_path)
                self.file_label.config(text=f"📄 {file_name}")
                
                self.log_message(f"✅ Excel 파일 로드 성공: {file_name}")
                self.log_message(f"📊 시트 수: {len(excel_data)}")
                
                # Excel 데이터를 테이블 형태로 변환
                self.extracted_tables = []
                for sheet_name, df in excel_data.items():
                    if not df.empty:
                        self.extracted_tables.append({
                            'name': sheet_name,
                            'data': df,
                            'rows': len(df),
                            'cols': len(df.columns)
                        })
                        self.log_message(f"  📋 {sheet_name}: {len(df)}행 x {len(df.columns)}열")
                
                self.update_table_display()
                
            except Exception as e:
                messagebox.showerror("오류", f"Excel 파일 읽기 실패: {str(e)}")
                self.log_message(f"❌ Excel 파일 읽기 실패: {str(e)}")
    
    def extract_tables(self):
        """HTML에서 테이블 추출"""
        if not self.html_content:
            messagebox.showwarning("경고", "먼저 HTML 파일을 열어주세요.")
            return
        
        self.log_message("\\n🔍 HTML 테이블 추출 시작...")
        
        try:
            soup = BeautifulSoup(self.html_content, 'html.parser')
            tables = soup.find_all('table')
            
            self.extracted_tables = []
            
            for i, table in enumerate(tables):
                # 테이블을 pandas DataFrame으로 변환
                try:
                    df = pd.read_html(str(table))[0]
                    
                    if len(df) > 1 and len(df.columns) > 1:  # 의미있는 크기의 테이블만
                        self.extracted_tables.append({
                            'name': f'Table_{i+1}',
                            'data': df,
                            'rows': len(df),
                            'cols': len(df.columns),
                            'html_table': table
                        })
                        
                        self.log_message(f"  📊 테이블 {i+1}: {len(df)}행 x {len(df.columns)}열")
                except:
                    continue
            
            self.log_message(f"✅ 총 {len(self.extracted_tables)}개 테이블 추출 완료")
            self.update_table_display()
            
        except Exception as e:
            self.log_message(f"❌ 테이블 추출 실패: {str(e)}")
    
    def update_table_display(self):
        """추출된 테이블을 Treeview에 표시"""
        # 기존 내용 삭제
        for item in self.table_tree.get_children():
            self.table_tree.delete(item)
        
        if not self.extracted_tables:
            return
        
        # 첫 번째 테이블의 컬럼 설정
        first_table = self.extracted_tables[0]['data']
        
        # 컬럼 설정
        columns = [f"Col_{i}" for i in range(len(first_table.columns))]
        self.table_tree['columns'] = columns
        self.table_tree['show'] = 'tree headings'
        
        # 컬럼 헤더 설정
        self.table_tree.heading('#0', text='테이블')
        for i, col in enumerate(columns):
            self.table_tree.heading(col, text=f"열{i+1}")
            self.table_tree.column(col, width=100)
        
        # 테이블 데이터 추가
        for table_info in self.extracted_tables:
            table_name = table_info['name']
            df = table_info['data']
            
            # 테이블 이름을 루트 노드로 추가
            table_node = self.table_tree.insert('', 'end', text=table_name, open=True)
            
            # 데이터 행 추가 (최대 10행까지만 표시)
            for idx, row in df.head(10).iterrows():
                values = [str(val)[:20] + ('...' if len(str(val)) > 20 else '') for val in row]
                self.table_tree.insert(table_node, 'end', text=f"행{idx+1}", values=values)
    
    def detect_levels(self):
        """재무제표 레벨 자동 감지"""
        if not self.extracted_tables:
            messagebox.showwarning("경고", "먼저 테이블을 추출해주세요.")
            return
        
        self.log_message("\\n📊 재무제표 레벨 자동 감지 시작...")
        
        for table_info in self.extracted_tables:
            df = table_info['data']
            table_name = table_info['name']
            
            self.log_message(f"\\n🔍 {table_name} 레벨 분석:")
            
            # 숫자 컬럼 찾기
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            # 각 행의 들여쓰기 레벨 감지 (첫 번째 컬럼 기준)
            if len(df.columns) > 0:
                first_col = df.iloc[:, 0].astype(str)
                
                levels = []
                for value in first_col:
                    # 공백 수로 레벨 판단
                    leading_spaces = len(value) - len(value.lstrip())
                    level = leading_spaces // 2  # 2칸 단위로 레벨 계산
                    levels.append(level)
                
                # 레벨 정보 저장
                table_info['levels'] = levels
                
                # 레벨별 통계
                level_counts = pd.Series(levels).value_counts().sort_index()
                self.log_message(f"  📈 감지된 레벨: {dict(level_counts)}")
                
                # 숫자 데이터가 있는 경우 레벨별 합계 분석
                if len(numeric_cols) > 0:
                    for col in numeric_cols[:3]:  # 처음 3개 숫자 컬럼만
                        self.analyze_level_sums(df, col, levels, table_name)
    
    def analyze_level_sums(self, df, column, levels, table_name):
        """레벨별 합계 분석"""
        try:
            numeric_data = pd.to_numeric(df[column], errors='coerce')
            
            for level in range(max(levels) + 1):
                level_mask = [l == level for l in levels]
                level_sum = numeric_data[level_mask].sum()
                level_count = sum(level_mask)
                
                if level_count > 0:
                    self.log_message(f"    Level {level} ({column}): {level_count}개 항목, 합계 {level_sum:,.0f}")
        
        except Exception as e:
            self.log_message(f"    ❌ {column} 레벨 합계 분석 실패: {str(e)}")
    
    def verify_sums(self):
        """합계 검증 수행"""
        if not self.extracted_tables:
            messagebox.showwarning("경고", "먼저 테이블을 추출해주세요.")
            return
        
        self.log_message("\\n➕ 합계 검증 시작...")
        
        verification_errors = []
        
        for table_info in self.extracted_tables:
            df = table_info['data']
            table_name = table_info['name']
            
            self.log_message(f"\\n🔍 {table_name} 합계 검증:")
            
            # 숫자 컬럼 찾기
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            for col in numeric_cols:
                try:
                    # 컬럼의 합계 계산
                    column_sum = df[col].sum()
                    
                    # '합계', '총계', 'Total' 등이 포함된 행 찾기
                    sum_keywords = ['합계', '총계', 'Total', '계', 'Sum']
                    
                    first_col = df.iloc[:, 0].astype(str)
                    
                    for idx, cell_value in first_col.items():
                        if any(keyword in str(cell_value) for keyword in sum_keywords):
                            reported_sum = df.loc[idx, col]
                            
                            if pd.notna(reported_sum) and abs(column_sum - reported_sum) > 0.01:
                                error_msg = f"합계 불일치: {col} 컬럼, 계산값 {column_sum:,.0f} ≠ 보고값 {reported_sum:,.0f}"
                                verification_errors.append(error_msg)
                                self.log_message(f"    ❌ {error_msg}")
                            else:
                                self.log_message(f"    ✅ {col}: 합계 일치 ({column_sum:,.0f})")
                            break
                
                except Exception as e:
                    self.log_message(f"    ⚠️ {col} 검증 실패: {str(e)}")
        
        if verification_errors:
            self.verification_results.extend(verification_errors)
            self.log_message(f"\\n⚠️ 총 {len(verification_errors)}개 합계 오류 발견")
        else:
            self.log_message("\\n✅ 모든 합계 검증 통과")
    
    def cross_reference_check(self):
        """교차 참조 확인"""
        if not self.extracted_tables:
            messagebox.showwarning("경고", "먼저 테이블을 추출해주세요.")
            return
        
        self.log_message("\\n🔄 교차 참조 확인 시작...")
        
        # 테이블 간 공통 항목 찾기
        if len(self.extracted_tables) >= 2:
            for i in range(len(self.extracted_tables)):
                for j in range(i + 1, len(self.extracted_tables)):
                    table1 = self.extracted_tables[i]
                    table2 = self.extracted_tables[j]
                    
                    self.compare_tables(table1, table2)
        else:
            self.log_message("  ⚠️ 교차 참조를 위해 최소 2개 테이블이 필요합니다")
    
    def compare_tables(self, table1, table2):
        """두 테이블 간 비교"""
        try:
            df1, df2 = table1['data'], table2['data']
            name1, name2 = table1['name'], table2['name']
            
            self.log_message(f"\\n🔍 {name1} ↔ {name2} 비교:")
            
            # 첫 번째 컬럼을 기준으로 공통 항목 찾기
            if len(df1.columns) > 0 and len(df2.columns) > 0:
                items1 = set(df1.iloc[:, 0].astype(str).str.strip())
                items2 = set(df2.iloc[:, 0].astype(str).str.strip())
                
                common_items = items1 & items2
                
                if common_items:
                    self.log_message(f"  📋 공통 항목 {len(common_items)}개 발견")
                    
                    # 공통 항목의 숫자 값 비교
                    self.compare_common_values(df1, df2, common_items, name1, name2)
                else:
                    self.log_message(f"  ℹ️ 공통 항목 없음")
        
        except Exception as e:
            self.log_message(f"  ❌ 테이블 비교 실패: {str(e)}")
    
    def compare_common_values(self, df1, df2, common_items, name1, name2):
        """공통 항목의 값 비교"""
        for item in list(common_items)[:5]:  # 처음 5개만 비교
            try:
                # 각 테이블에서 해당 항목의 행 찾기
                mask1 = df1.iloc[:, 0].astype(str).str.strip() == item
                mask2 = df2.iloc[:, 0].astype(str).str.strip() == item
                
                if mask1.any() and mask2.any():
                    row1 = df1[mask1].iloc[0]
                    row2 = df2[mask2].iloc[0]
                    
                    # 숫자 컬럼 비교
                    numeric_cols1 = df1.select_dtypes(include=[np.number]).columns
                    numeric_cols2 = df2.select_dtypes(include=[np.number]).columns
                    
                    for col1 in numeric_cols1:
                        for col2 in numeric_cols2:
                            val1 = row1[col1]
                            val2 = row2[col2]
                            
                            if pd.notna(val1) and pd.notna(val2) and abs(val1 - val2) < 0.01:
                                self.log_message(f"    ✅ {item}: {val1:,.0f} (일치)")
                            elif pd.notna(val1) and pd.notna(val2):
                                self.log_message(f"    ❌ {item}: {name1}={val1:,.0f} ≠ {name2}={val2:,.0f}")
            except:
                continue
    
    def detect_errors(self):
        """일반적인 오류 패턴 탐지"""
        if not self.extracted_tables:
            messagebox.showwarning("경고", "먼저 테이블을 추출해주세요.")
            return
        
        self.log_message("\\n⚠️ 오류 패턴 탐지 시작...")
        
        error_patterns = []
        
        for table_info in self.extracted_tables:
            df = table_info['data']
            table_name = table_info['name']
            
            self.log_message(f"\\n🔍 {table_name} 오류 검사:")
            
            # 1. 비정상적인 음수/양수 패턴
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            for col in numeric_cols:
                negative_count = (df[col] < 0).sum()
                positive_count = (df[col] > 0).sum()
                
                if negative_count > 0:
                    self.log_message(f"  📊 {col}: 음수 {negative_count}개, 양수 {positive_count}개")
                    
                    # 비정상적으로 많은 음수가 있는 경우
                    if negative_count > positive_count * 0.5:
                        error_patterns.append(f"{table_name}.{col}: 비정상적으로 많은 음수 값")
            
            # 2. 중복 항목 확인
            if len(df.columns) > 0:
                first_col = df.iloc[:, 0].astype(str)
                duplicates = first_col[first_col.duplicated()].unique()
                
                if len(duplicates) > 0:
                    self.log_message(f"  🔄 중복 항목 {len(duplicates)}개 발견")
                    for dup in duplicates[:3]:
                        error_patterns.append(f"{table_name}: 중복 항목 '{dup}'")
            
            # 3. 빈 셀이 많은 컬럼
            for col in df.columns:
                null_ratio = df[col].isnull().sum() / len(df)
                if null_ratio > 0.5:
                    error_patterns.append(f"{table_name}.{col}: 빈 셀 비율 {null_ratio:.1%}")
        
        if error_patterns:
            self.verification_results.extend(error_patterns)
            self.log_message(f"\\n⚠️ 총 {len(error_patterns)}개 오류 패턴 발견")
            for error in error_patterns:
                self.log_message(f"    ❌ {error}")
        else:
            self.log_message("\\n✅ 오류 패턴 없음")
    
    def generate_report(self):
        """검증 리포트 생성"""
        self.result_text.delete(1.0, tk.END)
        
        report = f"""
🔍 DSD Breaker 감사보고서 검증 리포트
{'='*60}

📁 파일: {os.path.basename(self.current_file) if self.current_file else '없음'}
📊 추출된 테이블: {len(self.extracted_tables)}개
⚠️ 발견된 문제: {len(self.verification_results)}개

📋 테이블 요약:
"""
        
        for i, table_info in enumerate(self.extracted_tables, 1):
            report += f"  {i}. {table_info['name']}: {table_info['rows']}행 x {table_info['cols']}열\\n"
        
        if self.verification_results:
            report += "\\n⚠️ 발견된 문제점:\\n"
            for i, result in enumerate(self.verification_results, 1):
                report += f"  {i}. {result}\\n"
        else:
            report += "\\n✅ 검증 결과: 문제 없음\\n"
        
        report += f"""
\\n📝 검증 권고사항:
• 자동 검증 결과를 참고하되, 반드시 수동 확인 필요
• 특히 복잡한 재무제표 구조는 별도 검토 권장
• 교차 참조가 불일치하는 항목은 원본 문서 재확인

🕒 검증 완료 시간: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
        
        self.result_text.insert(tk.END, report)
        self.notebook.select(2)  # 검증 결과 탭으로 이동
        
        self.log_message("\\n📋 검증 리포트 생성 완료")
    
    def run_full_verification(self):
        """전체 검증 프로세스 실행"""
        if not self.current_file:
            messagebox.showwarning("경고", "먼저 파일을 열어주세요.")
            return
        
        self.log_message("\\n🚀 전체 검증 프로세스 시작...")
        
        # 초기화
        self.verification_results = []
        
        # 순차적으로 모든 검증 실행
        if self.html_content and not self.extracted_tables:
            self.extract_tables()
        
        self.detect_levels()
        self.verify_sums()
        self.cross_reference_check()
        self.detect_errors()
        self.generate_report()
        
        self.log_message("\\n✅ 전체 검증 프로세스 완료")
    
    def save_results(self):
        """검증 결과 저장"""
        if not self.verification_results and not self.extracted_tables:
            messagebox.showwarning("경고", "저장할 결과가 없습니다.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="검증 결과 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("Text files", "*.txt")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.xlsx'):
                    # Excel 형태로 저장
                    with pd.ExcelWriter(file_path) as writer:
                        # 각 테이블을 별도 시트로 저장
                        for table_info in self.extracted_tables:
                            df = table_info['data']
                            sheet_name = table_info['name'][:31]  # Excel 시트명 길이 제한
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # 검증 결과 시트
                        if self.verification_results:
                            result_df = pd.DataFrame(self.verification_results, columns=['검증 결과'])
                            result_df.to_excel(writer, sheet_name='검증결과', index=False)
                else:
                    # 텍스트 형태로 저장
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(self.result_text.get(1.0, tk.END))
                
                self.log_message(f"✅ 결과 저장 완료: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror("오류", f"결과 저장 실패: {str(e)}")
                self.log_message(f"❌ 결과 저장 실패: {str(e)}")
    
    def show_help(self):
        """도움말 표시"""
        help_text = '''
🔍 DSD Breaker DART 감사보고서 검증 도구 사용법

1. 📁 파일 열기:
   • HTML 파일: DART에서 다운로드한 감사보고서 HTML
   • Excel 파일: 재무제표가 포함된 Excel 파일

2. 🛠️ 검증 기능:
   • 테이블 추출: HTML에서 재무 테이블 자동 추출
   • 레벨 감지: 재무제표 항목의 계층 구조 자동 감지
   • 합계 검증: 보고된 합계와 계산 합계 비교
   • 교차 참조: 여러 테이블 간 동일 항목 값 비교
   • 오류 탐지: 일반적인 오류 패턴 자동 탐지

3. 📋 결과 확인:
   • 처리 로그: 실시간 처리 과정 확인
   • 추출된 테이블: 파싱된 재무 데이터 확인
   • 검증 결과: 최종 검증 리포트

⚠️ 주의사항:
   • 자동 검증 결과는 참고용이며, 반드시 수동 확인 필요
   • 복잡한 재무제표 구조는 별도 검토 권장
        '''
        messagebox.showinfo("사용법", help_text)
    
    def show_about(self):
        """정보 표시"""
        about_text = '''
🔍 DSD Breaker DART 감사보고서 검증 도구 v2.0

DART (Data Analysis, Retrieval and Transfer) 
감사보고서의 자동 검증을 위한 Python 도구

✨ 주요 기능:
• HTML/Excel 파일 지원
• 재무제표 레벨 자동 감지
• 합계 검증 및 교차 참조
• 오류 패턴 자동 탐지
• 상세 검증 리포트 생성

🛠️ 개발: Python 3.x + pandas + BeautifulSoup
📅 업데이트: 2025년 6월 23일
        '''
        messagebox.showinfo("정보", about_text)
    
    def log_message(self, message):
        """로그 메시지 추가"""
        self.log_text.insert(tk.END, message + "\\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def run(self):
        """애플리케이션 실행"""
        self.root.mainloop()

def main():
    """메인 실행 함수"""
    app = DSDBreakAuditApp()
    app.run()

if __name__ == "__main__":
    main()