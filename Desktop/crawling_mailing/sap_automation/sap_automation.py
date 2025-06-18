from pywinauto.application import Application
import time
import logging
from datetime import datetime

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('sap_automation.log'),
        logging.StreamHandler()
    ]
)

class SAPAutomation:
    def __init__(self):
        self.app = None
        self.main_window = None
        
    def connect_to_sap(self):
        """SAP GUI에 연결"""
        try:
            # SAP GUI 프로세스에 연결
            self.app = Application(backend="win32").connect(path="saplgpad.exe")
            self.main_window = self.app.window(title_re=".*SAP.*")
            logging.info("SAP GUI에 성공적으로 연결되었습니다.")
            return True
        except Exception as e:
            logging.error(f"SAP GUI 연결 실패: {str(e)}")
            return False

    def login(self, client, user, password):
        """SAP 시스템에 로그인"""
        try:
            # 로그인 창 찾기
            login_window = self.app.window(title_re=".*SAP Logon.*")
            
            # 클라이언트 입력
            login_window.child_window(title="Client").type_keys(client)
            
            # 사용자 ID 입력
            login_window.child_window(title="User").type_keys(user)
            
            # 비밀번호 입력
            login_window.child_window(title="Password").type_keys(password)
            
            # 로그인 버튼 클릭
            login_window.child_window(title="Log On").click()
            
            logging.info("SAP 시스템에 성공적으로 로그인했습니다.")
            return True
        except Exception as e:
            logging.error(f"로그인 실패: {str(e)}")
            return False

    def run_transaction(self, transaction_code):
        """트랜잭션 실행"""
        try:
            # 트랜잭션 코드 입력 필드 찾기
            self.main_window.child_window(title="Command Field").type_keys(transaction_code)
            self.main_window.child_window(title="Command Field").type_keys("{ENTER}")
            logging.info(f"트랜잭션 {transaction_code} 실행 완료")
            return True
        except Exception as e:
            logging.error(f"트랜잭션 실행 실패: {str(e)}")
            return False

    def close(self):
        """SAP 연결 종료"""
        try:
            if self.main_window:
                self.main_window.close()
            logging.info("SAP 연결이 종료되었습니다.")
        except Exception as e:
            logging.error(f"SAP 종료 중 오류 발생: {str(e)}")

def main():
    # SAP 자동화 인스턴스 생성
    sap = SAPAutomation()
    
    try:
        # SAP GUI에 연결
        if not sap.connect_to_sap():
            return
        
        # SAP 시스템에 로그인
        if not sap.login(client="100", user="YOUR_USERNAME", password="YOUR_PASSWORD"):
            return
        
        # 여기에 자동화할 작업 추가
        # 예: 트랜잭션 실행
        sap.run_transaction("ME23N")
        
        # 작업 완료 후 잠시 대기
        time.sleep(5)
        
    except Exception as e:
        logging.error(f"작업 실행 중 오류 발생: {str(e)}")
    finally:
        # SAP 연결 종료
        sap.close()

if __name__ == "__main__":
    main() 