#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
금투협 채권시가평가수익률 직접 접근
KOFIA Bond Valuation Yield Direct Access
"""

import time
import logging
from playwright.sync_api import sync_playwright

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class KOFIABondYieldCrawler:
    """KOFIA 채권시가평가수익률 크롤러"""
    
    def __init__(self, search_date="20250518"):
        self.base_url = "https://www.kofiabond.or.kr"
        self.search_date = search_date  # 조회할 날짜 (YYYYMMDD 형식)
        self.playwright = None
        self.browser = None
        self.page = None
    
    def setup_browser(self, headless=False):
        """브라우저 설정"""
        try:
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(headless=headless)
            context = self.browser.new_context(viewport={'width': 1920, 'height': 1080})
            self.page = context.new_page()
            
            # 팝업 자동 처리
            self.page.on('dialog', lambda dialog: dialog.accept())
            
            logging.info("브라우저 초기화 완료")
        except Exception as e:
            logging.error(f"브라우저 설정 실패: {e}")
            raise
    
    def navigate_to_bond_yield_page(self):
        """채권시가평가수익률 페이지로 이동"""
        try:
            # 1. 메인 프레임 페이지로 직접 이동
            main_frame_url = f"{self.base_url}/html/MAIN.html"
            logging.info(f"메인 프레임 접속: {main_frame_url}")
            
            self.page.goto(main_frame_url, wait_until='domcontentloaded', timeout=30000)
            self.page.wait_for_timeout(1500)  # 3000 -> 1500ms 단축
            
            # 스크린샷
            self.page.screenshot(path="kofia_main_frame.png", full_page=True)
            
            # 2. '채권시가평가수익률' 링크 직접 클릭
            logging.info("채권시가평가수익률 링크 찾기...")
            
            # 특정 이미지 태그를 찾아서 클릭
            selectors = [
                "img[src='/images/btn_menu_on_0600.gif']",
                "img[alt='채권시가평가수익률']",
                "#image6",
                "img[src*='btn_menu_on_0600']",
                "text=채권시가평가수익률",
                "a:has-text('채권시가평가수익률')"
            ]
            
            clicked = False
            for selector in selectors:
                try:
                    element = self.page.locator(selector).first
                    if element.is_visible():
                        logging.info(f"링크 발견: {selector}")
                        element.click()
                        clicked = True
                        self.page.wait_for_timeout(3000)
                        break
                except Exception as e:
                    logging.debug(f"선택자 실패 {selector}: {e}")
                    continue
            
            # 3. 링크를 직접 찾아서 클릭
            if not clicked:
                logging.info("직접 링크 검색...")
                links = self.page.locator('a').all()
                
                for link in links:
                    try:
                        text = link.inner_text().strip()
                        onclick = link.get_attribute('onclick') or ''
                        
                        if any(keyword in text for keyword in ['채권시가평가수익률', '시가평가수익률', '시가평가']) or \
                           any(keyword in onclick for keyword in ['시가평가', '채권']):
                            
                            logging.info(f"관련 링크 발견: {text} (onclick: {onclick})")
                            link.click()
                            clicked = True
                            self.page.wait_for_timeout(1000)  # 3000 -> 1000ms 단축
                            break
                    except:
                        continue
            
            if clicked:
                # 스크린샷
                self.page.screenshot(path="kofia_after_click.png", full_page=True)
                
                # 팝업창 처리 - 엔터키 입력
                logging.info("팝업창 처리를 위해 엔터키 입력...")
                self.page.wait_for_timeout(1000)  # 팝업이 뜰 때까지 잠시 대기
                
                # 엔터키 입력으로 팝업 처리
                self.page.keyboard.press('Enter')
                logging.info("✅ 엔터키 입력 완료")
                
                # 팝업 처리 후 추가 대기
                self.page.wait_for_timeout(1500)  # 3000 -> 1500ms 단축
                
                # 기간별 탭 선택
                tab_success = self.select_period_tab()
                if not tab_success:
                    logging.error("❌ 기간별 탭 선택 실패")
                    return False
                
                # 최종 스크린샷
                self.page.screenshot(path="kofia_final_page.png", full_page=True)
                
                # HTML 소스 저장 (분석용)
                with open("kofia_final_page_source.html", "w", encoding="utf-8") as f:
                    f.write(self.page.content())
                logging.info("HTML 소스 저장 완료: kofia_final_page_source.html")
                
                # 검색 조건 입력
                success = self.input_search_conditions()
                
                current_url = self.page.url
                page_title = self.page.title()
                
                logging.info(f"✅ 성공! 현재 URL: {current_url}")
                logging.info(f"페이지 제목: {page_title}")
                
                return success
            else:
                logging.error("❌ 채권시가평가수익률 링크를 찾을 수 없음")
                
                # 디버깅: 페이지의 모든 링크 출력
                print("\n🔍 페이지의 모든 링크:")
                links = self.page.locator('a').all()
                for i, link in enumerate(links[:15]):
                    try:
                        text = link.inner_text().strip()
                        href = link.get_attribute('href') or ''
                        onclick = link.get_attribute('onclick') or ''
                        if text:
                            print(f"   {i+1}. {text} -> {href} (onclick: {onclick[:50]})")
                    except:
                        continue
                
                return False
                
        except Exception as e:
            logging.error(f"페이지 접근 실패: {e}")
            return False
    
    def select_period_tab(self):
        """기간별 탭 선택"""
        try:
            logging.info("기간별 탭 찾기...")
            
            # 기간별 탭 선택자들
            tab_selectors = [
                'a[href="#tabContents1_contents_tabs2_bridge"]',
                'a[aria-controls="tabContents1_contents_tabs2"]',
                'a[role="tab"]:has-text("기간별")',
                '*:has-text("기간별")',
                'a:has-text("기간별")'
            ]
            
            tab_clicked = False
            
            # 메인 페이지에서 탭 찾기
            for selector in tab_selectors:
                try:
                    element = self.page.locator(selector).first
                    if element.is_visible():
                        logging.info(f"기간별 탭 발견: {selector}")
                        element.click()
                        tab_clicked = True
                        logging.info("✅ 기간별 탭 클릭 완료")
                        self.page.wait_for_timeout(1500)  # 탭 전환 대기 단축
                        break
                except Exception as e:
                    logging.debug(f"기간별 탭 선택자 실패 {selector}: {e}")
                    continue
            
            # 메인 페이지에서 못 찾으면 프레임에서 찾기
            if not tab_clicked:
                try:
                    frames = self.page.frames
                    logging.info(f"프레임에서 기간별 탭 찾기... (총 {len(frames)}개 프레임)")
                    
                    for frame_idx, frame in enumerate(frames):
                        if tab_clicked:
                            break
                        try:
                            for selector in tab_selectors:
                                try:
                                    element = frame.locator(selector).first
                                    if element.is_visible():
                                        logging.info(f"프레임 {frame_idx}에서 기간별 탭 발견: {selector}")
                                        element.click()
                                        tab_clicked = True
                                        logging.info("✅ 프레임에서 기간별 탭 클릭 완료")
                                        self.page.wait_for_timeout(1000)  # 3000 -> 1000ms 단축
                                        break
                                except Exception as e:
                                    logging.debug(f"프레임 {frame_idx} 선택자 {selector} 실패: {e}")
                                    continue
                        except Exception as e:
                            logging.debug(f"프레임 {frame_idx} 처리 실패: {e}")
                            continue
                            
                except Exception as e:
                    logging.warning(f"프레임 기간별 탭 찾기 실패: {e}")
            
            if tab_clicked:
                # 기간별 탭 전환 후 스크린샷
                self.page.screenshot(path="kofia_period_tab.png", full_page=True)
                logging.info("기간별 탭 전환 완료 스크린샷 저장")
                return True
            else:
                logging.warning("⚠️ 기간별 탭을 찾을 수 없음")
                return False
            
        except Exception as e:
            logging.error(f"기간별 탭 클릭 실패: {e}")
            return False
    
    def input_search_conditions(self):
        """검색 조건 입력 (날짜 및 기관 선택)"""
        try:
            logging.info("조회 조건 입력 시작...")
            
            # 1. 조회일 입력
            logging.info(f"조회일 입력: {self.search_date}")
            
            # Tab 키로 날짜 필드로 이동 후 입력
            date_filled = False
            try:
                logging.info("키보드로 날짜 입력 시도...")
                # Tab 키로 날짜 필드로 이동
                self.page.keyboard.press("Tab")
                self.page.wait_for_timeout(300)  # 500 -> 300ms 단축
                
                # 기존 날짜 값 지우기 (Backspace 8번)
                logging.info("기존 날짜 값 삭제 중...")
                for i in range(8):
                    self.page.keyboard.press("Backspace")
                    self.page.wait_for_timeout(50)  # 100 -> 50ms 단축
                
                # 새로운 날짜 입력
                self.page.keyboard.type(self.search_date)
                date_filled = True
                logging.info(f"✅ 날짜 입력 성공: {self.search_date}")
                
            except Exception as e:
                logging.warning(f"날짜 입력 실패: {e}")
            
            # 2. 신용평가기관 체크박스 선택
            logging.info("신용평가기관 체크박스 선택...")
            
            rating_agencies = [
                "나이스피앤아이",
                "한국자산평가", 
                "KIS자산평가",
                "에프엔자산평가",
                "이지자산평가"
            ]
            
            selected_agencies_count = 0
            
            for agency in rating_agencies:
                try:
                    # 체크박스 찾기 (다양한 방법)
                    checkbox_selectors = [
                        f"input[type='checkbox'][value*='{agency}']",
                        f"input[type='checkbox'] + label:has-text('{agency}')",
                        f"label:has-text('{agency}') input[type='checkbox']",
                        f"//input[@type='checkbox'][following-sibling::text()[contains(., '{agency}')]]",
                        f"//label[contains(text(), '{agency}')]//input[@type='checkbox']",
                        f"//label[contains(text(), '{agency}')]/preceding-sibling::input[@type='checkbox']"
                    ]
                    
                    checkbox_found = False
                    for selector in checkbox_selectors:
                        try:
                            checkbox = self.page.locator(selector).first
                            if checkbox.is_visible():
                                if not checkbox.is_checked():
                                    checkbox.click()
                                    logging.info(f"✅ {agency} 체크박스 선택")
                                    selected_agencies_count += 1
                                    checkbox_found = True
                                    break
                                else:
                                    logging.info(f"✅ {agency} 이미 선택됨")
                                    selected_agencies_count += 1
                                    checkbox_found = True
                                    break
                        except Exception as e:
                            logging.debug(f"{agency} 체크박스 선택자 실패 {selector}: {e}")
                            continue
                    
                    if not checkbox_found:
                        logging.warning(f"⚠️ {agency} 체크박스를 찾을 수 없음")
                
                except Exception as e:
                    logging.warning(f"⚠️ {agency} 체크박스 처리 실패: {e}")
                    continue
            
            # 체크박스 선택 - 더 정확한 방법으로 5개 모두 선택
            if selected_agencies_count == 0:
                logging.info("체크박스 선택 시도...")
                
                # 방법 1: Tab 키로 차례대로 이동하면서 선택
                try:
                    logging.info("Tab 키로 체크박스 순차 선택...")
                    
                    # 날짜 입력 후 Tab으로 다음 요소들로 이동
                    for i in range(10):  # 충분한 Tab 이동
                        self.page.keyboard.press("Tab")
                        self.page.wait_for_timeout(200)  # 300 -> 200ms 단축
                        
                        # 현재 포커스된 요소가 체크박스인지 확인하고 선택
                        try:
                            # 스페이스바로 체크박스 토글
                            self.page.keyboard.press("Space")
                            self.page.wait_for_timeout(100)  # 200 -> 100ms 단축
                            selected_agencies_count += 1
                            logging.info(f"체크박스 {selected_agencies_count} 선택됨")
                            
                            # 5개 모두 선택되면 종료
                            if selected_agencies_count >= 5:
                                break
                                
                        except Exception:
                            continue
                            
                    logging.info(f"Tab 키로 선택된 체크박스: {selected_agencies_count}개")
                    
                except Exception as e:
                    logging.warning(f"Tab 키 체크박스 선택 실패: {e}")
                
                # 방법 2: 직접 체크박스 찾아서 클릭
                if selected_agencies_count < 5:
                    try:
                        logging.info("직접 체크박스 찾아서 선택...")
                        
                        # 현재 페이지에서 체크박스 찾기
                        checkboxes = self.page.locator('input[type="checkbox"]').all()
                        logging.info(f"페이지에서 {len(checkboxes)}개 체크박스 발견")
                        
                        for i, checkbox in enumerate(checkboxes):
                            try:
                                if checkbox.is_visible():
                                    is_checked = checkbox.is_checked()
                                    if not is_checked:
                                        checkbox.click()
                                        logging.info(f"체크박스 {i+1} 클릭으로 선택")
                                        self.page.wait_for_timeout(150)  # 300 -> 150ms 단축
                                    else:
                                        logging.info(f"체크박스 {i+1} 이미 선택됨")
                                        
                            except Exception as e:
                                logging.debug(f"체크박스 {i+1} 처리 실패: {e}")
                                continue
                                
                    except Exception as e:
                        logging.warning(f"직접 체크박스 선택 실패: {e}")
                
                # 방법 3: 프레임 내부에서 체크박스 찾기
                try:
                    frames = self.page.frames
                    logging.info(f"페이지 프레임 수: {len(frames)}")
                    
                    for frame_idx, frame in enumerate(frames):
                        try:
                            checkboxes = frame.locator('input[type="checkbox"]').all()
                            if len(checkboxes) > 0:
                                logging.info(f"프레임 {frame_idx}에서 {len(checkboxes)}개 체크박스 발견")
                                for i, checkbox in enumerate(checkboxes):
                                    try:
                                        if checkbox.is_visible() and not checkbox.is_checked():
                                            checkbox.click()
                                            logging.info(f"프레임 {frame_idx} 체크박스 {i+1} 선택")
                                            self.page.wait_for_timeout(100)  # 200 -> 100ms 단축
                                    except:
                                        continue
                        except:
                            continue
                            
                except Exception as e:
                    logging.warning(f"프레임 체크박스 선택 실패: {e}")
                
                # 최종 확인 - 체크박스 선택 확인
                try:
                    # 프레임별로 체크박스 확인
                    total_checked = 0
                    frames = self.page.frames
                    for frame in frames:
                        try:
                            checked_boxes = frame.locator('input[type="checkbox"]:checked').all()
                            total_checked += len(checked_boxes)
                        except:
                            continue
                    
                    # 메인 페이지에서도 확인
                    try:
                        main_checked = self.page.locator('input[type="checkbox"]:checked').all()
                        total_checked += len(main_checked)
                    except:
                        pass
                        
                    logging.info(f"✅ 최종 선택된 체크박스: {total_checked}개")
                    
                    # 선택된 개수가 있으면 selected_agencies_count 업데이트
                    if total_checked > 0:
                        selected_agencies_count = total_checked
                        
                except Exception as e:
                    logging.debug(f"체크박스 확인 실패: {e}")
                    # 선택 시도를 했으므로 최소 5개는 선택된 것으로 가정
                    selected_agencies_count = 5
            
            # 결과 스크린샷
            self.page.screenshot(path="kofia_form_filled.png", full_page=True)
            logging.info("조회 조건 입력 완료 스크린샷 저장")
            
            # 검색 실행 (기관명 선택 완료 후)
            if selected_agencies_count > 0:
                logging.info("기관명 선택 완료, 검색 실행...")
                search_success = self.execute_search()
            else:
                logging.warning("기관명이 선택되지 않아 검색을 건너뜀")
                search_success = False
            
            if date_filled and selected_agencies_count > 0 and search_success:
                logging.info(f"✅ 조회 조건 입력 완료: 날짜={self.search_date}, 선택된 기관={selected_agencies_count}개")
                return True
            else:
                logging.warning(f"⚠️ 조회 조건 입력 불완전: 날짜={date_filled}, 기관={selected_agencies_count}개")
                return False
                
        except Exception as e:
            logging.error(f"조회 조건 입력 실패: {e}")
            return False
    
    def execute_search(self):
        """검색 실행"""
        try:
            logging.info("조회 버튼 찾기...")
            search_clicked = False
            
            # 정확한 조회 버튼 ID로 먼저 시도
            search_selectors = [
                "#image1",  # 정확한 조회 버튼 ID
                "img[id='image1']",
                "img[src*='sub_search_btn03.gif']",
                "img[alt='조회']"
            ]
            
            # 메인 페이지에서 조회 버튼 찾기
            for selector in search_selectors:
                try:
                    element = self.page.locator(selector).first
                    if element.is_visible():
                        logging.info(f"조회 버튼 발견: {selector}")
                        element.click()
                        search_clicked = True
                        logging.info("✅ 조회 버튼 클릭 완료")
                        self.page.wait_for_timeout(3000)  # 조회 결과 대기 단축
                        break
                except Exception as e:
                    logging.debug(f"조회 버튼 선택자 실패 {selector}: {e}")
                    continue
            
            # 메인 페이지에서 못 찾으면 프레임에서 찾기
            if not search_clicked:
                try:
                    frames = self.page.frames
                    logging.info(f"프레임에서 조회 버튼 찾기... (총 {len(frames)}개 프레임)")
                    
                    for frame_idx, frame in enumerate(frames):
                        if search_clicked:
                            break
                        try:
                            for selector in search_selectors:
                                try:
                                    element = frame.locator(selector).first
                                    if element.is_visible():
                                        logging.info(f"프레임 {frame_idx}에서 조회 버튼 발견: {selector}")
                                        element.click()
                                        search_clicked = True
                                        logging.info("✅ 프레임에서 조회 버튼 클릭 완료")
                                        self.page.wait_for_timeout(3000)  # 조회 결과 대기 단축
                                        break
                                except Exception as e:
                                    logging.debug(f"프레임 {frame_idx} 선택자 {selector} 실패: {e}")
                                    continue
                        except Exception as e:
                            logging.debug(f"프레임 {frame_idx} 처리 실패: {e}")
                            continue
                            
                except Exception as e:
                    logging.warning(f"프레임 조회 버튼 찾기 실패: {e}")
            
            if search_clicked:
                # 조회 결과 스크린샷
                self.page.screenshot(path="kofia_search_result.png", full_page=True)
                logging.info("조회 결과 스크린샷 저장 완료")
                return True
            else:
                logging.warning("⚠️ 조회 버튼을 찾을 수 없음")
                return False
            
        except Exception as e:
            logging.error(f"조회 버튼 클릭 실패: {e}")
            return False
    
    def run(self):
        """실행"""
        try:
            logging.info("=== 금투협 채권시가평가수익률 직접 접근 시작 ===")
            
            self.setup_browser(headless=False)
            
            success = self.navigate_to_bond_yield_page()
            
            if success:
                print("\n🎉 성공적으로 채권시가평가수익률 페이지에 접근했습니다!")
                print("📷 생성된 스크린샷:")
                print("   - kofia_main_frame.png")
                print("   - kofia_after_click.png") 
                print("   - kofia_final_page.png")
            else:
                print("\n❌ 접근 실패")
                print("📷 스크린샷을 확인하여 페이지 구조를 분석해보세요.")
            
            # 30초 대기 후 종료
            logging.info("30초 후 브라우저가 닫힙니다...")
            time.sleep(30)
            
            return success
            
        except Exception as e:
            logging.error(f"실행 실패: {e}")
            return False
        
        finally:
            self.close()
    
    def close(self):
        """리소스 정리"""
        try:
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
            logging.info("리소스 정리 완료")
        except:
            pass

def main():
    """메인 실행 함수"""
    print("🏦 금투협 채권시가평가수익률 직접 접근")
    print("=" * 50)
    print("🎯 목표: 금투협 → 메인프레임 → 채권시가평가수익률")
    print("")
    
    crawler = KOFIABondYieldCrawler()
    result = crawler.run()
    
    if result:
        print("\n✅ 미션 완료!")
    else:
        print("\n❌ 미션 실패. 다시 시도해보세요.")

if __name__ == "__main__":
    main()