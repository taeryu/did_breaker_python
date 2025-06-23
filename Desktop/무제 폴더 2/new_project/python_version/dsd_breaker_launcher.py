#!/usr/bin/env python3
"""
🚀 DSD Breaker 통합 런처
일반 데이터 분석 / DART 감사보고서 검증 버전 선택 실행
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

class DSDBreakerLauncher:
    """DSD Breaker 버전 선택 런처"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🚀 DSD Breaker 런처")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        self.setup_ui()
    
    def setup_ui(self):
        """런처 UI 설정"""
        
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목
        title_label = ttk.Label(main_frame, text="🔧 DSD Breaker Python Edition", 
                               font=("Arial", 18, "bold"))
        title_label.pack(pady=(0, 10))
        
        subtitle_label = ttk.Label(main_frame, text="사용할 버전을 선택해주세요", 
                                  font=("Arial", 12))
        subtitle_label.pack(pady=(0, 30))
        
        # 버전 선택 프레임
        version_frame = ttk.Frame(main_frame)
        version_frame.pack(fill=tk.X, pady=(0, 20))
        
        # DSD Breaker 핵심 기능 (HTML → Excel 변환)
        converter_frame = ttk.LabelFrame(version_frame, text="🔄 DSD Breaker 핵심 기능 (HTML → Excel 변환)", padding="15")
        converter_frame.pack(fill=tk.X, pady=(0, 15))
        
        converter_desc = ttk.Label(converter_frame, text="""
✨ 핵심 기능:
• DART HTML 감사보고서 → Excel 파일 변환
• 복수 HTML 파일 일괄 처리
• 테이블별 시트 분리 옵션
• 데이터 정리 및 숫자 자동 인식
• 실시간 변환 진행 상황 표시

🎯 적합한 사용자:
• DART 감사보고서를 Excel로 변환해야 하는 사용자
• HTML 테이블 데이터를 Excel로 옮기고 싶은 사용자
• 원본 Excel Add-in 기능이 필요한 사용자
        """, justify=tk.LEFT)
        converter_desc.pack(anchor=tk.W)
        
        ttk.Button(converter_frame, text="🔄 HTML → Excel 변환기 실행", 
                  command=self.launch_converter_version,
                  style="Accent.TButton").pack(pady=(10, 0))
        
        # 일반 데이터 분석 버전
        general_frame = ttk.LabelFrame(version_frame, text="📊 일반 데이터 분석 버전", padding="15")
        general_frame.pack(fill=tk.X, pady=(0, 15))
        
        general_desc = ttk.Label(general_frame, text="""
✨ 주요 기능:
• Excel 파일 데이터 분석 및 시각화
• 6가지 차트 타입 (막대/선/산점도/히스토그램/파이/박스)
• 데이터 정리 및 변환
• 패턴 분석 및 상관관계 분석
• 한글 폰트 자동 설정

🎯 적합한 사용자:
• 일반적인 데이터 분석이 필요한 사용자
• Excel 데이터를 시각화하고 싶은 사용자
• 통계 분석 및 차트 생성이 필요한 사용자
        """, justify=tk.LEFT)
        general_desc.pack(anchor=tk.W)
        
        ttk.Button(general_frame, text="🚀 일반 분석 버전 실행", 
                  command=self.launch_general_version,
                  style="Accent.TButton").pack(pady=(10, 0))
        
        # DART 감사보고서 검증 버전
        audit_frame = ttk.LabelFrame(version_frame, text="🔍 DART 감사보고서 검증 버전", padding="15")
        audit_frame.pack(fill=tk.X)
        
        audit_desc = ttk.Label(audit_frame, text="""
✨ 주요 기능:
• DART HTML/Excel 감사보고서 자동 검증
• 재무제표 레벨 자동 감지 (들여쓰기 기반)
• 합계 검증 (보고값 vs 계산값 비교)
• 교차 참조 확인 (테이블 간 일관성 검증)
• 오류 패턴 자동 탐지

🎯 적합한 사용자:
• 감사 업무 담당자
• 재무제표 검증이 필요한 회계사
• DART 보고서 분석 담당자
        """, justify=tk.LEFT)
        audit_desc.pack(anchor=tk.W)
        
        ttk.Button(audit_frame, text="🔍 감사 검증 버전 실행", 
                  command=self.launch_audit_version,
                  style="Accent.TButton").pack(pady=(10, 0))
        
        # 하단 정보
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(20, 0))
        
        info_text = ttk.Label(info_frame, text="""
📋 추가 정보:
• 두 버전 모두 독립적으로 실행 가능
• 필요에 따라 여러 버전을 동시에 사용 가능
• 모든 기능은 한글을 완벽 지원

💡 팁: 처음 사용하시는 경우 테스트 데이터를 먼저 생성해보세요.
        """, justify=tk.LEFT, font=("Arial", 9))
        info_text.pack()
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Button(button_frame, text="📊 일반용 테스트 데이터 생성", 
                  command=self.create_general_test_data).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="🔍 감사용 테스트 데이터 생성", 
                  command=self.create_audit_test_data).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="❌ 종료", 
                  command=self.root.quit).pack(side=tk.RIGHT)
    
    def check_dependencies(self):
        """의존성 체크"""
        required_modules = ['pandas', 'numpy', 'matplotlib', 'openpyxl']
        missing_modules = []
        
        for module in required_modules:
            try:
                __import__(module)
            except ImportError:
                missing_modules.append(module)
        
        if missing_modules:
            messagebox.showerror("의존성 오류", 
                f"다음 모듈이 설치되지 않았습니다:\\n{', '.join(missing_modules)}\\n\\n설치 명령어:\\npip install {' '.join(missing_modules)}")
            return False
        
        return True
    
    def launch_converter_version(self):
        """HTML → Excel 변환기 실행"""
        if not self.check_dependencies():
            return
        
        # xlsxwriter 체크
        try:
            import xlsxwriter
        except ImportError:
            messagebox.showerror("의존성 오류", 
                "HTML → Excel 변환에는 xlsxwriter가 필요합니다:\\npip install xlsxwriter")
            return
        
        try:
            from dsd_breaker_converter import DSDHTMLToExcelConverter
            
            # 현재 창 숨기기
            self.root.withdraw()
            
            # 변환기 실행
            app = DSDHTMLToExcelConverter()
            app.run()
            
            # 실행 완료 후 런처 창 다시 표시
            self.root.deiconify()
            
        except ImportError as e:
            messagebox.showerror("실행 오류", f"HTML → Excel 변환기를 찾을 수 없습니다:\\n{e}")
        except Exception as e:
            messagebox.showerror("실행 오류", f"예상치 못한 오류:\\n{e}")
            self.root.deiconify()
    
    def launch_general_version(self):
        """일반 데이터 분석 버전 실행"""
        if not self.check_dependencies():
            return
        
        try:
            from dsd_breaker_concept import DSDBreakerApp
            
            # 현재 창 숨기기
            self.root.withdraw()
            
            # 일반 버전 실행
            app = DSDBreakerApp()
            app.run()
            
            # 실행 완료 후 런처 창 다시 표시
            self.root.deiconify()
            
        except ImportError as e:
            messagebox.showerror("실행 오류", f"일반 분석 버전을 찾을 수 없습니다:\\n{e}")
        except Exception as e:
            messagebox.showerror("실행 오류", f"예상치 못한 오류:\\n{e}")
            self.root.deiconify()
    
    def launch_audit_version(self):
        """DART 감사보고서 검증 버전 실행"""
        if not self.check_dependencies():
            return
        
        # BeautifulSoup 체크
        try:
            import bs4
        except ImportError:
            messagebox.showerror("의존성 오류", 
                "감사 검증 버전에는 beautifulsoup4가 필요합니다:\\npip install beautifulsoup4")
            return
        
        try:
            from dsd_breaker_audit import DSDBreakAuditApp
            
            # 현재 창 숨기기
            self.root.withdraw()
            
            # 감사 버전 실행
            app = DSDBreakAuditApp()
            app.run()
            
            # 실행 완료 후 런처 창 다시 표시
            self.root.deiconify()
            
        except ImportError as e:
            messagebox.showerror("실행 오류", f"감사 검증 버전을 찾을 수 없습니다:\\n{e}")
        except Exception as e:
            messagebox.showerror("실행 오류", f"예상치 못한 오류:\\n{e}")
            self.root.deiconify()
    
    def create_general_test_data(self):
        """일반용 테스트 데이터 생성"""
        try:
            import subprocess
            result = subprocess.run([sys.executable, "test_dsd_breaker.py"], 
                                  capture_output=True, text=True)
            
            if result.returncode == 0:
                messagebox.showinfo("성공", "일반용 테스트 데이터가 생성되었습니다:\\n• sample_data.xlsx")
            else:
                messagebox.showerror("오류", f"테스트 데이터 생성 실패:\\n{result.stderr}")
                
        except Exception as e:
            messagebox.showerror("오류", f"테스트 데이터 생성 실패:\\n{e}")
    
    def create_audit_test_data(self):
        """감사용 테스트 데이터 생성"""
        try:
            import subprocess
            result = subprocess.run([sys.executable, "test_audit_features.py"], 
                                  capture_output=True, text=True)
            
            if result.returncode == 0:
                messagebox.showinfo("성공", 
                    "감사용 테스트 데이터가 생성되었습니다:\\n• sample_financial_statements.xlsx\\n• sample_audit_report.html")
            else:
                messagebox.showerror("오류", f"테스트 데이터 생성 실패:\\n{result.stderr}")
                
        except Exception as e:
            messagebox.showerror("오류", f"테스트 데이터 생성 실패:\\n{e}")
    
    def run(self):
        """런처 실행"""
        self.root.mainloop()

def main():
    """메인 실행 함수"""
    print("🚀 DSD Breaker 통합 런처 시작")
    
    launcher = DSDBreakerLauncher()
    launcher.run()
    
    print("👋 DSD Breaker 런처 종료")

if __name__ == "__main__":
    main()