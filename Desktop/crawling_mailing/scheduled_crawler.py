import pickle
import schedule
import time
import pandas as pd
import smtplib
import os
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import undetected_chromedriver as uc
from datetime import datetime
from email.message import EmailMessage

# ✅ 크롤링할 URL
URL = "https://www.bandtrass.or.kr/customs/total.do?command=CUS001View&viewCode=CUS00401"

# ✅ 네이버 메일 SMTP 설정
SMTP_SERVER = "smtp.naver.com"
SMTP_PORT = 587
EMAIL_SENDER = "gpt821225@naver.com"  # ✅ 네이버 메일 주소 입력
EMAIL_PASSWORD = "game11.."  # ✅ 네이버 비밀번호 (앱 비밀번호 사용)
EMAIL_RECEIVER = "prwtaeryu@gmail.com"  # ✅ 받는 사람 이메일 입력

def run_crawler():
    """📌 크롤링 실행 및 엑셀 저장"""
    print(f"🚀 {datetime.now()} - 크롤링 시작!")

    # ✅ Chrome 브라우저 설정
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")  # ✅ 자동화 감지 방지
    options.add_argument("start-maximized")  # ✅ 브라우저 전체 화면으로 실행
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(options=options)

    try:
        # ✅ 웹페이지 열기
        driver.get(URL)
        time.sleep(5)  # 페이지 로딩 대기

        # ✅ 저장된 쿠키 불러오기
        with open("cookies.pkl", "rb") as f:
            cookies = pickle.load(f)

        for cookie in cookies:
            driver.add_cookie(cookie)

        driver.refresh()  # ✅ 쿠키 적용 후 새로고침
        time.sleep(5)

        # ✅ 품목 코드 입력 필드 찾기
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#SelectCd"))
        )
        search_box.clear()
        search_box.send_keys("3005904000")  # ✅ 품목 코드 입력
        time.sleep(2)
        search_box.send_keys("\n")   # ✅ 엔터 키 입력 (조회 실행)

        print("✅ 품목 코드 입력 & 검색 완료! 테이블 로딩 대기 중...")

        # ✅ 테이블이 로딩될 시간을 주기 위해 10초 대기
        time.sleep(10)

        # ✅ BeautifulSoup으로 페이지 파싱
        soup = BeautifulSoup(driver.page_source, "html.parser")

        # ✅ 테이블 요소 찾기 (모든 행 `tr` 선택)
        table_rows = soup.select("#table_list_1 > tbody > tr")

        if not table_rows:
            print("❌ 테이블을 찾을 수 없습니다! `driver.page_source`를 확인하세요.")
            with open("error_page.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)  # 🔥 HTML 저장해서 분석 가능하게 함.
            return None  # 크롤링 실패하면 종료

        # ✅ 데이터 추출
        data = []
        for row in table_rows:
            cells = [cell.text.strip() for cell in row.find_all("td")]
            if cells:
                data.append(cells)

        # ✅ Pandas DataFrame 생성
        df = pd.DataFrame(data)

        # ✅ 날짜별 파일명 생성 후 저장
        today = datetime.now().strftime("%Y-%m-%d")
        file_name = f"관세청_수출입_데이터_{today}.xlsx"
        df.to_excel(file_name, index=False)
        print(f"✅ 엑셀 파일 저장 완료! ({file_name})")

        return file_name  # 크롤링 완료 후 파일명 반환

    finally:
        driver.quit()  # ✅ 웹드라이버 종료

def find_latest_excel():
    """📌 가장 최신 엑셀 파일을 찾아서 '피상회복제_수출.xlsx'로 저장"""
    files = glob.glob("관세청_수출입_데이터_*.xlsx")  # ✅ 날짜별 저장된 엑셀 파일 찾기
    if not files:
        print("❌ 엑셀 파일이 존재하지 않습니다. 메일 전송 취소!")
        return None

    latest_file = max(files, key=os.path.getctime)  # ✅ 최신 파일 찾기
    new_file_path = "피상회복제_수출.xlsx"

    # ✅ 최신 파일을 "피상회복제_수출.xlsx"로 복사
    os.rename(latest_file, new_file_path)
    print(f"✅ 최신 파일 선택: {latest_file} -> {new_file_path}")
    return new_file_path

def send_email():
    """📌 최신 엑셀 파일을 찾아서 메일로 전송"""
    file_path = find_latest_excel()

    if not file_path:
        return  # 파일이 없으면 메일 전송 취소

    # ✅ 이메일 메시지 생성
    msg = EmailMessage()
    msg["Subject"] = "📊 피상회복제 수출 데이터 자동 발송"
    msg["From"] = EMAIL_SENDER
    msg["To"] = EMAIL_RECEIVER
    msg.set_content("안녕하세요,\n\n자동으로 크롤링된 피상회복제 수출 데이터를 첨부합니다.\n\n감사합니다.")

    # ✅ 파일 첨부
    with open(file_path, "rb") as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_path)

    try:
        # ✅ SMTP 서버 연결
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # 보안 연결
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)  # 로그인
            server.send_message(msg)  # 이메일 전송
        print("✅ 이메일 전송 완료!")

    except Exception as e:
        print(f"❌ 이메일 전송 실패: {e}")

def check_and_run():
    """📌 특정 날짜(1일, 11일, 21일)에만 실행"""
    today = datetime.now()
    if today.day in [1, 11, 21] :
        file_name = run_crawler()  # ✅ 크롤링 실행
        if file_name:  # ✅ 크롤링 성공하면 메일 전송 실행
            send_email()

# ✅ 매일 정오(12:00)에 실행, 날짜를 확인 후 실행 여부 결정
schedule.every().day.at("12:00").do(check_and_run)

print("⏳ 크롤링 자동 실행 대기 중... (매월 1일, 11일, 21일 정오 실행)")

while True:
    schedule.run_pending()
    time.sleep(60)  # 1분마다 체크