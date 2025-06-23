#!/usr/bin/env python3
"""
DSD Breaker Python 버전 컨셉
기존 Excel Add-in을 Python으로 재구현
"""

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np

# seaborn을 선택적으로 임포트
try:
    import seaborn as sns
    HAS_SEABORN = True
except ImportError:
    HAS_SEABORN = False

class DSDBreakerApp:
    """DSD Breaker Python 버전 메인 애플리케이션"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DSD Breaker - Python Edition")
        self.root.geometry("800x600")
        
        # 상태 변수
        self.current_file = None
        self.data = None
        self.chart_window = None
        
        # 한글 폰트 설정
        self.setup_korean_font()
        
        self.setup_ui()
    
    def setup_korean_font(self):
        """한글 폰트 설정 (matplotlib용)"""
        try:
            # macOS 기본 한글 폰트 설정
            if os.name == 'posix':  # macOS/Linux
                font_candidates = ['AppleGothic', 'Malgun Gothic', 'NanumGothic']
            else:  # Windows
                font_candidates = ['Malgun Gothic', 'NanumGothic', 'Gulim']
            
            for font_name in font_candidates:
                try:
                    plt.rcParams['font.family'] = font_name
                    plt.rcParams['axes.unicode_minus'] = False
                    break
                except:
                    continue
            
            print(f"✅ 한글 폰트 설정 완료")
            
        except Exception as e:
            print(f"⚠️ 한글 폰트 설정 실패: {str(e)}")
        
    def setup_ui(self):
        """사용자 인터페이스 설정"""
        
        # 메뉴바
        self.create_menubar()
        
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="🔧 DSD Breaker", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 파일 선택 영역
        file_frame = ttk.LabelFrame(main_frame, text="📁 파일 선택", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(file_frame, text="Excel 파일 열기", 
                  command=self.open_file).grid(row=0, column=0, padx=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="파일을 선택해주세요")
        self.file_label.grid(row=0, column=1, sticky=tk.W)
        
        # 기능 버튼 영역
        functions_frame = ttk.LabelFrame(main_frame, text="🛠️ 데이터 처리 기능", padding="10")
        functions_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 기능 버튼들 (예상되는 기능들)
        buttons = [
            ("📊 데이터 분석", self.analyze_data),
            ("🔄 데이터 변환", self.transform_data),
            ("📈 차트 생성", self.create_chart),
            ("💾 결과 저장", self.save_results),
            ("🧹 데이터 정리", self.clean_data),
            ("🔍 패턴 찾기", self.find_patterns)
        ]
        
        for i, (text, command) in enumerate(buttons):
            row = i // 3
            col = i % 3
            ttk.Button(functions_frame, text=text, command=command, 
                      width=15).grid(row=row, column=col, padx=5, pady=5)
        
        # 결과 표시 영역
        results_frame = ttk.LabelFrame(main_frame, text="📋 결과", padding="10")
        results_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 텍스트 영역 (스크롤 포함)
        text_frame = ttk.Frame(results_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.result_text = tk.Text(text_frame, height=15, width=80)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        # 초기 메시지
        self.log_message("🎉 DSD Breaker Python Edition에 오신 것을 환영합니다!")
        self.log_message("📁 Excel 파일을 열어서 시작해보세요.")
        
    def create_menubar(self):
        """메뉴바 생성"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 파일 메뉴
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="열기", command=self.open_file)
        file_menu.add_command(label="저장", command=self.save_results)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.root.quit)
        
        # 도구 메뉴
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도구", menu=tools_menu)
        tools_menu.add_command(label="데이터 분석", command=self.analyze_data)
        tools_menu.add_command(label="데이터 정리", command=self.clean_data)
        
        # 도움말 메뉴
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도움말", menu=help_menu)
        help_menu.add_command(label="사용법", command=self.show_help)
        help_menu.add_command(label="정보", command=self.show_about)
    
    def open_file(self):
        """Excel 파일 열기"""
        file_path = filedialog.askopenfilename(
            title="Excel 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.current_file = file_path
                self.data = pd.read_excel(file_path)
                
                file_name = os.path.basename(file_path)
                self.file_label.config(text=f"📄 {file_name}")
                
                self.log_message(f"✅ 파일 로드 성공: {file_name}")
                self.log_message(f"📊 데이터 크기: {self.data.shape[0]}행 x {self.data.shape[1]}열")
                self.log_message(f"📋 컬럼: {', '.join(self.data.columns.tolist())}")
                
            except Exception as e:
                messagebox.showerror("오류", f"파일 읽기 실패: {str(e)}")
                self.log_message(f"❌ 파일 읽기 실패: {str(e)}")
    
    def analyze_data(self):
        """데이터 분석"""
        if self.data is None:
            messagebox.showwarning("경고", "먼저 Excel 파일을 열어주세요.")
            return
        
        self.log_message("\\n🔍 데이터 분석 시작...")
        
        # 기본 통계
        self.log_message(f"📊 기본 정보:")
        self.log_message(f"  - 총 행 수: {len(self.data):,}")
        self.log_message(f"  - 총 열 수: {len(self.data.columns)}")
        self.log_message(f"  - 메모리 사용량: {self.data.memory_usage(deep=True).sum() / 1024**2:.2f} MB")
        
        # 누락 데이터
        missing_data = self.data.isnull().sum()
        if missing_data.any():
            self.log_message(f"\\n⚠️ 누락 데이터:")
            for col, count in missing_data[missing_data > 0].items():
                self.log_message(f"  - {col}: {count}개")
        else:
            self.log_message("\\n✅ 누락 데이터 없음")
        
        # 수치 데이터 요약
        numeric_cols = self.data.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            self.log_message(f"\\n📈 수치 데이터 요약:")
            stats = self.data[numeric_cols].describe()
            self.log_message(stats.to_string())
    
    def transform_data(self):
        """데이터 변환"""
        if self.data is None:
            messagebox.showwarning("경고", "먼저 Excel 파일을 열어주세요.")
            return
        
        self.log_message("\\n🔄 데이터 변환 시작...")
        
        # 예시: 간단한 데이터 변환들
        try:
            # 1. 컬럼명 정리
            original_cols = self.data.columns.tolist()
            self.data.columns = [col.strip().replace(' ', '_') for col in self.data.columns]
            self.log_message("✅ 컬럼명 정리 완료 (공백 제거, 언더스코어 변환)")
            
            # 2. 중복 행 제거
            before_count = len(self.data)
            self.data = self.data.drop_duplicates()
            after_count = len(self.data)
            removed = before_count - after_count
            
            if removed > 0:
                self.log_message(f"✅ 중복 행 {removed}개 제거")
            else:
                self.log_message("✅ 중복 행 없음")
            
            # 3. 데이터 타입 최적화
            self.log_message("🔧 데이터 타입 최적화 중...")
            for col in self.data.columns:
                if self.data[col].dtype == 'object':
                    try:
                        # 숫자로 변환 가능한지 확인
                        pd.to_numeric(self.data[col], errors='raise')
                        self.data[col] = pd.to_numeric(self.data[col])
                        self.log_message(f"  📊 {col}: 숫자형으로 변환")
                    except:
                        pass
            
            self.log_message("✅ 데이터 변환 완료!")
            
        except Exception as e:
            self.log_message(f"❌ 데이터 변환 실패: {str(e)}")
    
    def create_chart(self):
        """차트 생성"""
        if self.data is None:
            messagebox.showwarning("경고", "먼저 Excel 파일을 열어주세요.")
            return
        
        self.log_message("\\n📈 차트 생성 시작...")
        
        # 새 창에서 차트 옵션 선택
        self.show_chart_options()
    
    def show_chart_options(self):
        """차트 옵션 선택 창 표시"""
        if self.chart_window and self.chart_window.winfo_exists():
            self.chart_window.destroy()
        
        self.chart_window = tk.Toplevel(self.root)
        self.chart_window.title("📈 차트 생성 옵션")
        self.chart_window.geometry("600x500")
        
        # 메인 프레임
        main_frame = ttk.Frame(self.chart_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목
        title_label = ttk.Label(main_frame, text="📊 차트 생성 옵션", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # 수치 컬럼 선택
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns.tolist()
        
        if not numeric_cols:
            ttk.Label(main_frame, text="❌ 수치 데이터가 없어 차트를 생성할 수 없습니다.").pack()
            return
        
        # 컬럼 선택 프레임
        col_frame = ttk.LabelFrame(main_frame, text="📋 데이터 컬럼 선택", padding="10")
        col_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(col_frame, text="X축 컬럼:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.x_column = ttk.Combobox(col_frame, values=self.data.columns.tolist(), width=20)
        self.x_column.grid(row=0, column=1, padx=(0, 20))
        self.x_column.set(self.data.columns[0])
        
        ttk.Label(col_frame, text="Y축 컬럼:").grid(row=0, column=2, sticky=tk.W, padx=(0, 10))
        self.y_column = ttk.Combobox(col_frame, values=numeric_cols, width=20)
        self.y_column.grid(row=0, column=3)
        if numeric_cols:
            self.y_column.set(numeric_cols[0])
        
        # 차트 타입 선택
        chart_frame = ttk.LabelFrame(main_frame, text="📈 차트 타입", padding="10")
        chart_frame.pack(fill=tk.X, pady=(0, 10))
        
        chart_types = [
            ("📊 막대 차트", "bar"),
            ("📈 선 그래프", "line"),
            ("🔴 산점도", "scatter"),
            ("📉 히스토그램", "hist"),
            ("🥧 파이 차트", "pie"),
            ("📦 박스 플롯", "box")
        ]
        
        self.chart_type = tk.StringVar(value="bar")
        
        for i, (text, value) in enumerate(chart_types):
            row = i // 3
            col = i % 3
            ttk.Radiobutton(chart_frame, text=text, variable=self.chart_type, 
                           value=value).grid(row=row, column=col, sticky=tk.W, padx=10, pady=5)
        
        # 차트 옵션
        options_frame = ttk.LabelFrame(main_frame, text="⚙️ 차트 옵션", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.show_grid = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="격자 표시", variable=self.show_grid).grid(row=0, column=0, sticky=tk.W)
        
        self.chart_title = tk.StringVar(value="데이터 분석 차트")
        ttk.Label(options_frame, text="제목:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        ttk.Entry(options_frame, textvariable=self.chart_title, width=30).grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=(10, 0))
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Button(button_frame, text="📈 차트 생성", 
                  command=self.generate_chart).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="❌ 취소", 
                  command=self.chart_window.destroy).pack(side=tk.LEFT)
    
    def generate_chart(self):
        """실제 차트 생성"""
        try:
            chart_type = self.chart_type.get()
            x_col = self.x_column.get()
            y_col = self.y_column.get()
            title = self.chart_title.get()
            
            # 새 창에 차트 표시
            chart_display = tk.Toplevel(self.root)
            chart_display.title(f"📊 {title}")
            chart_display.geometry("800x600")
            
            # matplotlib 피규어 생성
            fig, ax = plt.subplots(figsize=(10, 6))
            
            if chart_type == "bar":
                # 막대 차트
                if self.data[x_col].dtype == 'object':
                    value_counts = self.data[x_col].value_counts().head(10)
                    ax.bar(value_counts.index, value_counts.values)
                    ax.set_xlabel(x_col)
                    ax.set_ylabel("빈도수")
                else:
                    ax.bar(range(len(self.data[x_col][:20])), self.data[y_col][:20])
                    ax.set_xlabel(x_col)
                    ax.set_ylabel(y_col)
                    
            elif chart_type == "line":
                # 선 그래프
                if len(self.data) > 1000:
                    sample_data = self.data.sample(1000).sort_index()
                else:
                    sample_data = self.data
                ax.plot(sample_data[x_col], sample_data[y_col])
                ax.set_xlabel(x_col)
                ax.set_ylabel(y_col)
                
            elif chart_type == "scatter":
                # 산점도
                if len(self.data) > 1000:
                    sample_data = self.data.sample(1000)
                else:
                    sample_data = self.data
                ax.scatter(sample_data[x_col], sample_data[y_col], alpha=0.6)
                ax.set_xlabel(x_col)
                ax.set_ylabel(y_col)
                
            elif chart_type == "hist":
                # 히스토그램
                ax.hist(self.data[y_col].dropna(), bins=30, alpha=0.7)
                ax.set_xlabel(y_col)
                ax.set_ylabel("빈도수")
                
            elif chart_type == "pie":
                # 파이 차트
                if self.data[x_col].dtype == 'object':
                    value_counts = self.data[x_col].value_counts().head(8)
                    ax.pie(value_counts.values, labels=value_counts.index, autopct='%1.1f%%')
                else:
                    self.log_message("❌ 파이 차트는 범주형 데이터가 필요합니다.")
                    return
                    
            elif chart_type == "box":
                # 박스 플롯
                numeric_cols = self.data.select_dtypes(include=[np.number]).columns[:5]
                ax.boxplot([self.data[col].dropna() for col in numeric_cols])
                ax.set_xticklabels(numeric_cols, rotation=45)
            
            # 차트 옵션 적용
            ax.set_title(title, fontsize=14, fontweight='bold')
            if self.show_grid.get():
                ax.grid(True, alpha=0.3)
            
            plt.tight_layout()
            
            # tkinter에 차트 임베드
            canvas = FigureCanvasTkAgg(fig, chart_display)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            # 저장 버튼 추가
            save_frame = ttk.Frame(chart_display)
            save_frame.pack(fill=tk.X, padx=10, pady=5)
            
            ttk.Button(save_frame, text="💾 차트 저장", 
                      command=lambda: self.save_chart(fig)).pack(side=tk.LEFT)
            ttk.Button(save_frame, text="🔄 새로고침", 
                      command=lambda: self.generate_chart()).pack(side=tk.LEFT, padx=(10, 0))
            
            self.log_message(f"✅ {chart_type} 차트 생성 완료")
            
            # 차트 옵션 창 닫기
            if self.chart_window:
                self.chart_window.destroy()
                
        except Exception as e:
            self.log_message(f"❌ 차트 생성 실패: {str(e)}")
            messagebox.showerror("오류", f"차트 생성 실패: {str(e)}")
    
    def save_chart(self, fig):
        """차트를 파일로 저장"""
        file_path = filedialog.asksaveasfilename(
            title="차트 저장",
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), ("SVG files", "*.svg")]
        )
        
        if file_path:
            try:
                fig.savefig(file_path, dpi=300, bbox_inches='tight')
                self.log_message(f"✅ 차트 저장 완료: {os.path.basename(file_path)}")
            except Exception as e:
                self.log_message(f"❌ 차트 저장 실패: {str(e)}")
    
    def save_results(self):
        """결과 저장"""
        if self.data is None:
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="결과 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.data.to_csv(file_path, index=False, encoding='utf-8-sig')
                else:
                    self.data.to_excel(file_path, index=False)
                
                self.log_message(f"✅ 파일 저장 완료: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror("오류", f"파일 저장 실패: {str(e)}")
                self.log_message(f"❌ 파일 저장 실패: {str(e)}")
    
    def clean_data(self):
        """데이터 정리"""
        if self.data is None:
            messagebox.showwarning("경고", "먼저 Excel 파일을 열어주세요.")
            return
        
        self.log_message("\\n🧹 데이터 정리 시작...")
        
        try:
            original_shape = self.data.shape
            
            # 1. 빈 행/열 제거
            self.data = self.data.dropna(how='all')  # 모든 값이 NaN인 행 제거
            self.data = self.data.loc[:, ~self.data.isnull().all()]  # 모든 값이 NaN인 열 제거
            
            # 2. 공백 문자 정리
            for col in self.data.select_dtypes(include=['object']).columns:
                self.data[col] = self.data[col].astype(str).str.strip()
                self.data[col] = self.data[col].replace('nan', None)
            
            new_shape = self.data.shape
            
            self.log_message(f"✅ 데이터 정리 완료:")
            self.log_message(f"  이전: {original_shape[0]}행 x {original_shape[1]}열")
            self.log_message(f"  이후: {new_shape[0]}행 x {new_shape[1]}열")
            
        except Exception as e:
            self.log_message(f"❌ 데이터 정리 실패: {str(e)}")
    
    def find_patterns(self):
        """패턴 찾기"""
        if self.data is None:
            messagebox.showwarning("경고", "먼저 Excel 파일을 열어주세요.")
            return
        
        self.log_message("\\n🔍 패턴 분석 시작...")
        
        try:
            # 1. 데이터 타입별 분포
            self.log_message("📊 데이터 타입 분포:")
            type_counts = self.data.dtypes.value_counts()
            for dtype, count in type_counts.items():
                self.log_message(f"  {dtype}: {count}개 컬럼")
            
            # 2. 수치 데이터 상관관계
            numeric_cols = self.data.select_dtypes(include=['number']).columns
            if len(numeric_cols) > 1:
                self.log_message("\\n📈 수치 데이터 상관관계 (상위 5개):")
                corr_matrix = self.data[numeric_cols].corr()
                
                # 대각선 제거하고 상관계수가 높은 쌍 찾기
                import numpy as np
                mask = np.triu(np.ones_like(corr_matrix, dtype=bool))
                corr_matrix = corr_matrix.mask(mask)
                
                # 절댓값 기준으로 정렬
                corr_pairs = []
                for i in range(len(corr_matrix.columns)):
                    for j in range(len(corr_matrix.columns)):
                        if not np.isnan(corr_matrix.iloc[i, j]):
                            corr_pairs.append((
                                corr_matrix.columns[i],
                                corr_matrix.columns[j],
                                corr_matrix.iloc[i, j]
                            ))
                
                corr_pairs.sort(key=lambda x: abs(x[2]), reverse=True)
                
                for col1, col2, corr in corr_pairs[:5]:
                    self.log_message(f"  {col1} ↔ {col2}: {corr:.3f}")
            
            self.log_message("✅ 패턴 분석 완료!")
            
        except Exception as e:
            self.log_message(f"❌ 패턴 분석 실패: {str(e)}")
    
    def show_help(self):
        """도움말 표시"""
        help_text = '''
🔧 DSD Breaker Python Edition 사용법

1. 📁 파일 열기: Excel 파일(.xlsx, .xls)을 선택합니다.

2. 🛠️ 데이터 처리 기능:
   - 📊 데이터 분석: 기본 통계와 데이터 정보를 표시
   - 🔄 데이터 변환: 컬럼명 정리, 중복 제거, 타입 최적화
   - 📈 차트 생성: 막대/선/산점도/히스토그램/파이/박스 차트 생성
   - 💾 결과 저장: 처리된 데이터를 Excel/CSV로 저장
   - 🧹 데이터 정리: 빈 행/열 제거, 공백 정리
   - 🔍 패턴 찾기: 데이터 패턴과 상관관계 분석

3. 📈 차트 기능:
   - 6가지 차트 타입 지원 (막대, 선, 산점도, 히스토그램, 파이, 박스)
   - 한글 폰트 자동 설정
   - PNG/PDF/SVG 형식으로 저장 가능
   - 실시간 차트 옵션 설정

4. 📋 결과: 모든 처리 과정과 결과가 하단에 표시됩니다.
        '''
        messagebox.showinfo("사용법", help_text)
    
    def show_about(self):
        """정보 표시"""
        about_text = '''
🔧 DSD Breaker Python Edition v1.0

Excel Add-in을 Python으로 재구현한 데이터 처리 도구

개발: Python 3.x + tkinter + pandas
목적: 효율적인 데이터 분석 및 처리
        '''
        messagebox.showinfo("정보", about_text)
    
    def log_message(self, message):
        """결과 영역에 메시지 로깅"""
        self.result_text.insert(tk.END, message + "\\n")
        self.result_text.see(tk.END)  # 자동 스크롤
        self.root.update()  # UI 업데이트
    
    def run(self):
        """애플리케이션 실행"""
        self.root.mainloop()

def main():
    """메인 실행 함수"""
    app = DSDBreakerApp()
    app.run()

if __name__ == "__main__":
    main()