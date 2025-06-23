#!/usr/bin/env python3
"""
🔄 DSD Breaker HTML → Excel 변환기
DART HTML 감사보고서를 Excel 파일로 변환하는 핵심 기능
"""

import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
import re
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
import xlsxwriter

class DSDHTMLToExcelConverter:
    """DART HTML을 Excel로 변환하는 DSD Breaker 핵심 기능"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🔄 DSD Breaker - HTML → Excel 변환기")
        self.root.geometry("900x700")
        
        # 상태 변수
        self.html_files = []
        self.converted_tables = []
        self.output_path = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """사용자 인터페이스 설정"""
        
        # 메뉴바
        self.create_menubar()
        
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="🔄 DSD Breaker - DART HTML → Excel 변환기", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 설명
        desc_label = ttk.Label(main_frame, 
                              text="DART에서 다운로드한 HTML 감사보고서를 Excel 파일로 변환합니다", 
                              font=("Arial", 10))
        desc_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # 파일 선택 영역
        file_frame = ttk.LabelFrame(main_frame, text="📁 HTML 파일 선택", padding="10")
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 파일 선택 버튼들
        btn_frame = ttk.Frame(file_frame)
        btn_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        ttk.Button(btn_frame, text="📄 HTML 파일 선택", 
                  command=self.select_html_files).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(btn_frame, text="📂 폴더 선택", 
                  command=self.select_html_folder).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(btn_frame, text="🗑️ 목록 지우기", 
                  command=self.clear_file_list).grid(row=0, column=2)
        
        # 선택된 파일 목록
        ttk.Label(file_frame, text="선택된 HTML 파일:").grid(row=1, column=0, sticky=tk.W, pady=(10, 5))
        
        # 파일 목록 표시
        list_frame = ttk.Frame(file_frame)
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.file_listbox = tk.Listbox(list_frame, height=4)
        list_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=list_scrollbar.set)
        
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        list_frame.columnconfigure(0, weight=1)
        
        # 변환 옵션
        options_frame = ttk.LabelFrame(main_frame, text="⚙️ 변환 옵션", padding="10")
        options_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 옵션 체크박스들
        self.preserve_formatting = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="원본 서식 유지", 
                       variable=self.preserve_formatting).grid(row=0, column=0, sticky=tk.W)
        
        self.split_by_table = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="테이블별 시트 분리", 
                       variable=self.split_by_table).grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        
        self.clean_data = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="데이터 정리 (공백 제거)", 
                       variable=self.clean_data).grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        
        self.detect_numbers = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="숫자 자동 인식", 
                       variable=self.detect_numbers).grid(row=1, column=1, sticky=tk.W, padx=(20, 0), pady=(5, 0))
        
        # 변환 실행 버튼
        convert_frame = ttk.Frame(main_frame)
        convert_frame.grid(row=4, column=0, columnspan=3, pady=(10, 0))
        
        ttk.Button(convert_frame, text="🔄 Excel로 변환", 
                  command=self.convert_to_excel,
                  style="Accent.TButton").grid(row=0, column=0, padx=(0, 10))
        ttk.Button(convert_frame, text="📊 미리보기", 
                  command=self.preview_conversion).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(convert_frame, text="📂 출력 폴더 열기", 
                  command=self.open_output_folder).grid(row=0, column=2)
        
        # 진행 상황 표시
        progress_frame = ttk.LabelFrame(main_frame, text="📋 변환 진행 상황", padding="10")
        progress_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 진행 바
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 로그 텍스트 영역
        self.log_text = scrolledtext.ScrolledText(progress_frame, height=15, width=80)
        self.log_text.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)
        file_frame.columnconfigure(0, weight=1)
        
        # 초기 메시지
        self.log_message("🎉 DSD Breaker HTML → Excel 변환기를 시작합니다")
        self.log_message("📁 DART에서 다운로드한 HTML 파일을 선택하세요")
    
    def create_menubar(self):
        """메뉴바 생성"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 파일 메뉴
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="HTML 파일 선택", command=self.select_html_files)
        file_menu.add_command(label="폴더 선택", command=self.select_html_folder)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.root.quit)
        
        # 변환 메뉴
        convert_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="변환", menu=convert_menu)
        convert_menu.add_command(label="Excel로 변환", command=self.convert_to_excel)
        convert_menu.add_command(label="미리보기", command=self.preview_conversion)
        
        # 도움말 메뉴
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도움말", menu=help_menu)
        help_menu.add_command(label="사용법", command=self.show_help)
        help_menu.add_command(label="정보", command=self.show_about)
    
    def select_html_files(self):
        """HTML 파일들 선택"""
        files = filedialog.askopenfilenames(
            title="DART HTML 파일 선택",
            filetypes=[("HTML files", "*.html *.htm"), ("All files", "*.*")]
        )
        
        if files:
            self.html_files.extend(files)
            self.update_file_list()
            self.log_message(f"✅ {len(files)}개 HTML 파일 추가됨")
    
    def select_html_folder(self):
        """HTML 파일이 있는 폴더 선택"""
        folder = filedialog.askdirectory(title="DART HTML 파일 폴더 선택")
        
        if folder:
            html_files = []
            for ext in ['*.html', '*.htm']:
                html_files.extend(Path(folder).glob(ext))
            
            if html_files:
                self.html_files.extend([str(f) for f in html_files])
                self.update_file_list()
                self.log_message(f"✅ 폴더에서 {len(html_files)}개 HTML 파일 발견")
            else:
                messagebox.showwarning("경고", "선택한 폴더에 HTML 파일이 없습니다.")
    
    def clear_file_list(self):
        """파일 목록 지우기"""
        self.html_files = []
        self.update_file_list()
        self.log_message("🗑️ 파일 목록이 지워졌습니다")
    
    def update_file_list(self):
        """파일 목록 업데이트"""
        self.file_listbox.delete(0, tk.END)
        for file_path in self.html_files:
            file_name = os.path.basename(file_path)
            self.file_listbox.insert(tk.END, file_name)
    
    def convert_to_excel(self):
        """HTML을 Excel로 변환"""
        if not self.html_files:
            messagebox.showwarning("경고", "변환할 HTML 파일을 선택해주세요.")
            return
        
        # 출력 파일 경로 선택
        output_file = filedialog.asksaveasfilename(
            title="Excel 파일 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not output_file:
            return
        
        self.output_path = output_file
        
        self.log_message(f"\\n🔄 변환 시작: {len(self.html_files)}개 파일")
        self.log_message(f"📂 출력 파일: {os.path.basename(output_file)}")
        
        try:
            # Excel 워크북 생성
            workbook = xlsxwriter.Workbook(output_file)
            
            # 스타일 정의
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D9E1F2',
                'border': 1,
                'align': 'center'
            })
            
            data_format = workbook.add_format({
                'border': 1,
                'align': 'left'
            })
            
            number_format = workbook.add_format({
                'border': 1,
                'align': 'right',
                'num_format': '#,##0'
            })
            
            total_files = len(self.html_files)
            
            for i, html_file in enumerate(self.html_files):
                self.progress_var.set((i / total_files) * 100)
                self.root.update()
                
                file_name = os.path.basename(html_file)
                self.log_message(f"\\n📄 처리 중: {file_name}")
                
                try:
                    # HTML 파일 읽기
                    with open(html_file, 'r', encoding='utf-8') as f:
                        html_content = f.read()
                    
                    # HTML 파싱 및 테이블 추출
                    tables = self.extract_tables_from_html(html_content)
                    
                    if not tables:
                        self.log_message(f"  ⚠️ {file_name}에서 테이블을 찾을 수 없습니다")
                        continue
                    
                    self.log_message(f"  📊 {len(tables)}개 테이블 발견")
                    
                    # 테이블을 Excel 시트로 변환
                    if self.split_by_table.get():
                        # 테이블별로 별도 시트 생성
                        for j, table in enumerate(tables):
                            sheet_name = f"{file_name[:20]}_Table{j+1}"
                            # Excel 시트명 규칙에 맞게 조정
                            sheet_name = re.sub(r'[\\\\/*?:\\\\[\\\\]\\n]', '_', sheet_name)[:31]
                            
                            worksheet = workbook.add_worksheet(sheet_name)
                            self.write_table_to_worksheet(worksheet, table, header_format, data_format, number_format)
                            
                            self.log_message(f"    ✅ 시트 생성: {sheet_name}")
                    else:
                        # 파일별로 하나의 시트에 모든 테이블
                        sheet_name = file_name[:31].replace('.html', '').replace('.htm', '')
                        sheet_name = re.sub(r'[\\\\/*?:\\\\[\\\\]\\n]', '_', sheet_name)
                        
                        worksheet = workbook.add_worksheet(sheet_name)
                        
                        row_offset = 0
                        for j, table in enumerate(tables):
                            if j > 0:
                                row_offset += 2  # 테이블 사이 간격
                            
                            rows_written = self.write_table_to_worksheet(
                                worksheet, table, header_format, data_format, number_format, row_offset)
                            row_offset += rows_written
                        
                        self.log_message(f"    ✅ 시트 생성: {sheet_name}")
                    
                except Exception as e:
                    self.log_message(f"  ❌ {file_name} 처리 실패: {str(e)}")
                    continue
            
            # 워크북 저장
            workbook.close()
            
            self.progress_var.set(100)
            self.log_message(f"\\n🎉 변환 완료!")
            self.log_message(f"📂 저장 위치: {output_file}")
            
            # 변환 완료 다이얼로그
            result = messagebox.askyesno("변환 완료", 
                f"Excel 파일 변환이 완료되었습니다!\\n\\n파일을 열어보시겠습니까?\\n\\n{os.path.basename(output_file)}")
            
            if result:
                self.open_output_file()
                
        except Exception as e:
            self.log_message(f"❌ 변환 실패: {str(e)}")
            messagebox.showerror("오류", f"변환 중 오류가 발생했습니다:\\n{str(e)}")
    
    def extract_tables_from_html(self, html_content):
        """HTML에서 테이블 추출"""
        soup = BeautifulSoup(html_content, 'html.parser')
        tables = soup.find_all('table')
        
        extracted_tables = []
        
        for table in tables:
            try:
                # pandas read_html 사용
                df_list = pd.read_html(str(table))
                
                if df_list:
                    df = df_list[0]
                    
                    # 데이터 정리 옵션 적용
                    if self.clean_data.get():
                        df = self.clean_dataframe(df)
                    
                    # 숫자 인식 옵션 적용
                    if self.detect_numbers.get():
                        df = self.convert_numbers(df)
                    
                    # 의미있는 크기의 테이블만 추가
                    if len(df) > 0 and len(df.columns) > 0:
                        extracted_tables.append(df)
                        
            except Exception as e:
                # pandas로 읽기 실패시 직접 파싱 시도
                try:
                    df = self.parse_table_manually(table)
                    if df is not None and len(df) > 0:
                        extracted_tables.append(df)
                except:
                    continue
        
        return extracted_tables
    
    def clean_dataframe(self, df):
        """데이터프레임 정리"""
        # 문자열 컬럼의 공백 제거
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
            # 'nan' 문자열을 실제 NaN으로 변환
            df[col] = df[col].replace('nan', np.nan)
        
        return df
    
    def convert_numbers(self, df):
        """숫자 자동 인식 및 변환"""
        for col in df.columns:
            if df[col].dtype == 'object':
                # 숫자 패턴 확인 (콤마 포함)
                try:
                    # 콤마 제거 후 숫자 변환 시도
                    cleaned_series = df[col].astype(str).str.replace(',', '').str.replace('\\s+', '', regex=True)
                    
                    # 숫자로 변환 가능한지 확인
                    numeric_series = pd.to_numeric(cleaned_series, errors='coerce')
                    
                    # 50% 이상이 숫자로 변환 가능하면 적용
                    if numeric_series.notna().sum() / len(df) > 0.5:
                        df[col] = numeric_series
                        
                except:
                    continue
        
        return df
    
    def parse_table_manually(self, table):
        """수동으로 테이블 파싱"""
        rows = table.find_all('tr')
        if not rows:
            return None
        
        data = []
        for row in rows:
            cells = row.find_all(['td', 'th'])
            row_data = [cell.get_text(strip=True) for cell in cells]
            if row_data:  # 빈 행 제외
                data.append(row_data)
        
        if not data:
            return None
        
        # 최대 컬럼 수 찾기
        max_cols = max(len(row) for row in data)
        
        # 모든 행을 같은 길이로 맞추기
        for row in data:
            while len(row) < max_cols:
                row.append('')
        
        # DataFrame 생성
        df = pd.DataFrame(data[1:], columns=data[0] if len(data) > 1 else None)
        
        return df
    
    def write_table_to_worksheet(self, worksheet, df, header_format, data_format, number_format, start_row=0):
        """테이블을 워크시트에 쓰기"""
        current_row = start_row
        
        # 헤더 쓰기
        for col_idx, col_name in enumerate(df.columns):
            worksheet.write(current_row, col_idx, str(col_name), header_format)
        
        current_row += 1
        
        # 데이터 쓰기
        for row_idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                # 숫자인지 확인하여 적절한 형식 적용
                if pd.notna(value) and isinstance(value, (int, float)):
                    worksheet.write(current_row, col_idx, value, number_format)
                else:
                    worksheet.write(current_row, col_idx, str(value) if pd.notna(value) else '', data_format)
            
            current_row += 1
        
        # 컬럼 너비 자동 조정
        for col_idx, col_name in enumerate(df.columns):
            max_width = max(
                len(str(col_name)),
                df.iloc[:, col_idx].astype(str).str.len().max() if len(df) > 0 else 0
            )
            worksheet.set_column(col_idx, col_idx, min(max_width + 2, 50))
        
        return current_row - start_row
    
    def preview_conversion(self):
        """변환 미리보기"""
        if not self.html_files:
            messagebox.showwarning("경고", "미리보기할 HTML 파일을 선택해주세요.")
            return
        
        # 첫 번째 파일로 미리보기
        first_file = self.html_files[0]
        
        try:
            with open(first_file, 'r', encoding='utf-8') as f:
                html_content = f.read()
            
            tables = self.extract_tables_from_html(html_content)
            
            if not tables:
                messagebox.showinfo("미리보기", "선택한 파일에서 테이블을 찾을 수 없습니다.")
                return
            
            # 미리보기 창 생성
            preview_window = tk.Toplevel(self.root)
            preview_window.title(f"📊 미리보기 - {os.path.basename(first_file)}")
            preview_window.geometry("800x600")
            
            # 노트북 위젯으로 테이블별 탭 생성
            notebook = ttk.Notebook(preview_window)
            notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            for i, table in enumerate(tables[:5]):  # 처음 5개 테이블만 미리보기
                tab_frame = ttk.Frame(notebook)
                notebook.add(tab_frame, text=f"테이블 {i+1}")
                
                # 테이블 표시용 텍스트 위젯
                text_widget = scrolledtext.ScrolledText(tab_frame, wrap=tk.NONE, font=("Courier", 9))
                text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
                
                # DataFrame을 문자열로 변환하여 표시
                table_str = table.head(20).to_string(index=False)  # 처음 20행만
                text_widget.insert(tk.END, table_str)
                text_widget.config(state=tk.DISABLED)
            
            self.log_message(f"📊 미리보기 표시: {len(tables)}개 테이블 (최대 5개)")
            
        except Exception as e:
            messagebox.showerror("오류", f"미리보기 실패:\\n{str(e)}")
    
    def open_output_folder(self):
        """출력 폴더 열기"""
        if self.output_path and os.path.exists(self.output_path):
            folder_path = os.path.dirname(self.output_path)
            
            # 운영체제별로 폴더 열기
            import subprocess
            import platform
            
            system = platform.system()
            try:
                if system == "Windows":
                    subprocess.run(["explorer", folder_path])
                elif system == "Darwin":  # macOS
                    subprocess.run(["open", folder_path])
                else:  # Linux
                    subprocess.run(["xdg-open", folder_path])
                    
                self.log_message(f"📂 출력 폴더 열기: {folder_path}")
            except:
                messagebox.showinfo("정보", f"출력 폴더: {folder_path}")
        else:
            messagebox.showwarning("경고", "아직 변환된 파일이 없습니다.")
    
    def open_output_file(self):
        """출력 파일 열기"""
        if self.output_path and os.path.exists(self.output_path):
            import subprocess
            import platform
            
            system = platform.system()
            try:
                if system == "Windows":
                    subprocess.run(["start", self.output_path], shell=True)
                elif system == "Darwin":  # macOS
                    subprocess.run(["open", self.output_path])
                else:  # Linux
                    subprocess.run(["xdg-open", self.output_path])
                    
                self.log_message(f"📂 Excel 파일 열기: {os.path.basename(self.output_path)}")
            except Exception as e:
                messagebox.showerror("오류", f"파일 열기 실패:\\n{str(e)}")
    
    def show_help(self):
        """도움말 표시"""
        help_text = '''
🔄 DSD Breaker HTML → Excel 변환기 사용법

1. 📁 HTML 파일 선택:
   • "HTML 파일 선택": 개별 파일 선택
   • "폴더 선택": 폴더 내 모든 HTML 파일 선택
   • 여러 파일을 선택하여 한 번에 변환 가능

2. ⚙️ 변환 옵션:
   • 원본 서식 유지: HTML 테이블의 기본 서식 보존
   • 테이블별 시트 분리: 각 테이블을 별도 시트로 생성
   • 데이터 정리: 불필요한 공백 제거
   • 숫자 자동 인식: 텍스트 형태의 숫자를 자동 변환

3. 🔄 변환 실행:
   • "Excel로 변환": 실제 변환 수행
   • "미리보기": 변환 결과 미리 확인
   • "출력 폴더 열기": 변환된 파일 위치 확인

💡 팁:
   • DART에서 다운로드한 HTML 파일에 최적화됨
   • 복잡한 테이블 구조도 자동으로 처리
   • 큰 파일의 경우 변환 시간이 소요될 수 있음
        '''
        messagebox.showinfo("사용법", help_text)
    
    def show_about(self):
        """정보 표시"""
        about_text = '''
🔄 DSD Breaker HTML → Excel 변환기 v2.0

DART(Data Analysis, Retrieval and Transfer) 
HTML 감사보고서를 Excel 파일로 변환하는 핵심 도구

✨ 주요 기능:
• HTML 테이블 자동 추출 및 Excel 변환
• 복수 파일 일괄 처리
• 다양한 변환 옵션 제공
• 실시간 진행 상황 표시

🛠️ 개발: Python 3.x + BeautifulSoup + pandas + xlsxwriter
📅 업데이트: 2025년 6월 23일

💡 원본 Excel Add-in의 핵심 기능을 Python으로 재구현
        '''
        messagebox.showinfo("정보", about_text)
    
    def log_message(self, message):
        """로그 메시지 추가"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\\n"
        
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.root.update()
    
    def run(self):
        """애플리케이션 실행"""
        self.root.mainloop()

def main():
    """메인 실행 함수"""
    app = DSDHTMLToExcelConverter()
    app.run()

if __name__ == "__main__":
    main()