import pandas as pd
import time
from datetime import datetime
from playwright.sync_api import sync_playwright
import logging
import os

class CandoAutoCounseling:
    def __init__(self, csv_file_path, login_url="https://cando.hoseo.ac.kr/Office/Home.aspx"):
        """
        호서대학교 Cando 시스템 자동 상담 입력 클래스
        
        Args:
            csv_file_path (str): CSV 파일 경로
            login_url (str): 로그인 URL
        """
        self.csv_file_path = csv_file_path
        self.login_url = login_url
        self.playwright = None  # 추가
        self.browser = None
        self.page = None
        self.results = []
        
        # 로깅 설정
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('cando_auto_log.txt', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
        # 상수 정의
        self.STATUS_MAPPING = {'일반': '1', '관심': '2', '중점': '3'}
        self.STATUS_FIELDS = {
            '진로상태': 'P', '취업상태': 'J', 
            '학습상태': 'C', '심리상태': 'M'
        }
        
    def load_csv_data(self):
        """CSV 파일을 로드하고 검증"""
        try:
            self.df = pd.read_csv(self.csv_file_path, encoding='utf-8')
            self.logger.info(f"CSV 파일 로드 완료: {len(self.df)}건의 상담 데이터")
            
            # 필수 컬럼 확인
            required_columns = ['학번', '이름', '상담일자', '상담내용']
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            
            if missing_columns:
                self.logger.error(f"필수 컬럼이 없습니다: {missing_columns}")
                return False
            
            # 누락 데이터 경고
            for col in required_columns:
                if self.df[col].isnull().any():
                    self.logger.warning(f"필수 컬럼 '{col}'에 누락된 데이터가 있습니다.")
            
            return True
            
        except Exception as e:
            self.logger.error(f"CSV 파일 로드 실패: {str(e)}")
            return False
    
    def setup_browser(self):  # self 추가!
        """브라우저 설정 - 시스템 Chrome 우선, 실패시 Playwright Chromium"""
        try:
            self.playwright = sync_playwright().start()
            
            # 방법 1: 시스템 Chrome 시도
            try:
                self.logger.info("시스템 Chrome 브라우저 연결 시도...")
                self.browser = self.playwright.chromium.launch(
                    headless=False,
                    channel="chrome",
                    args=[
                        '--start-maximized',
                        '--disable-blink-features=AutomationControlled'
                    ]
                )
                self.logger.info("✅ 시스템 Chrome 브라우저 사용")
            except Exception as chrome_error:
                # 방법 2: Playwright Chromium 사용
                self.logger.warning(f"시스템 Chrome 연결 실패: {chrome_error}")
                self.logger.info("Playwright Chromium 브라우저 시도...")
                
                try:
                    self.browser = self.playwright.chromium.launch(
                        headless=False,
                        args=['--start-maximized']
                    )
                    self.logger.info("✅ Playwright Chromium 브라우저 사용")
                except Exception as chromium_error:
                    self.logger.error("Playwright 브라우저도 없습니다.")
                    raise Exception(
                        "브라우저를 찾을 수 없습니다.\n"
                        "1. Chrome 브라우저를 설치하거나\n"
                        "2. 명령 프롬프트에서 'playwright install chromium' 실행"
                    )
            
            # 컨텍스트 및 페이지 생성
            context = self.browser.new_context(
                viewport={'width': 1920, 'height': 1080},
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            )
            
            self.page = context.new_page()
            self.logger.info("브라우저 설정 완료")
            return True
            
        except Exception as e:
            self.logger.error(f"브라우저 설정 실패: {e}")
            self.cleanup()
            return False
    
    def wait_for_login(self):
        """로그인 대기 및 확인"""
        try:
            self.logger.info("로그인 페이지로 이동 중...")
            self.page.goto(self.login_url)
            
            print("\n" + "="*60)
            print("🔐 호서대학교 Cando 시스템 자동 상담 입력 프로그램")
            print("="*60)
            print("📋 1단계: 로그인 대기")
            print("💡 브라우저에서 수동으로 로그인을 진행해주세요.")
            print("✅ 로그인 완료 후 아래 메시지가 나타나면 엔터키를 눌러주세요.")
            print("-"*60)
            
            input("🎯 로그인 완료 후 엔터키를 눌러주세요: ")
            
            # 로그인 확인
            try:
                self.page.wait_for_selector("h3:has-text('내 지도학생')", timeout=10000)
                self.logger.info("✅ 로그인 성공 확인")
                print("✅ 로그인이 성공적으로 확인되었습니다!")
                return True
            except:
                self.logger.error("❌ 로그인 확인 실패")
                print("❌ 로그인 확인에 실패했습니다. 다시 시도해주세요.")
                return False
                
        except Exception as e:
            self.logger.error(f"로그인 대기 중 오류: {str(e)}")
            return False
    
    def search_student(self, student_id):
        """학생 검색"""
        try:
            search_input = self.page.get_by_role("textbox", name="이름/학번")
            search_input.clear()
            search_input.fill(str(student_id))
            
            self.page.get_by_text("조회", exact=True).click()
            time.sleep(2)
            
            # 검색 결과 확인
            if str(student_id) in self.page.text_content("body"):
                self.logger.info(f"✅ 학생 검색 성공: {student_id}")
                return True
            else:
                self.logger.warning(f"⚠️ 학생 검색 실패: {student_id}")
                return False
                
        except Exception as e:
            self.logger.error(f"학생 검색 중 오류 ({student_id}): {str(e)}")
            return False

    def open_student_profile(self, student_id):
        """학생 프로필 열기"""
        try:
            student_listitem = self.page.get_by_role("listitem").filter(has_text=student_id)
            if student_listitem.is_visible():
                student_listitem.click()
                time.sleep(3)
                
                self.page.wait_for_selector("iframe", timeout=10000)
                self.logger.info(f"✅ 학생 프로필 열기 성공: {student_id}")
                return True
            else:
                self.logger.error(f"❌ 학생 항목을 찾을 수 없습니다: {student_id}")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ 학생 프로필 열기 실패 ({student_id}): {str(e)}")
            return False

    def input_counseling_data(self, row_data):
        """상담 데이터 입력"""
        try:
            iframe = self.page.frame_locator("iframe")
            
            # 상담입력 버튼 클릭
            self.logger.info("🟡 상담입력 버튼 클릭 (goCounsel)")
            iframe.locator('div[onclick="goCounsel()"]').click()
            
            # 상담입력 폼 대기
            self.logger.info("🟡 상담입력 폼 대기")
            iframe.locator("#Pdate").wait_for(state="visible", timeout=10000)
            
            # 기본 정보 입력
            self._input_basic_info(iframe, row_data)
            
            # 상담 내용 입력
            self._input_counseling_content(iframe, row_data)
            
            # 학생상태 입력 (기존 코드 유지)
            self._input_student_status(iframe, row_data)
            
            # 전문상담의뢰 입력 (기존 코드 유지)
            self._input_referral_status(iframe, row_data)
            
            # 비공개설정 (기존 코드 유지)
            self._set_privacy_option(iframe, row_data)
            
            self.logger.info(f"✅ 상담 데이터 입력 완료: {row_data['학번']}")
            return True
                    
        except Exception as e:
            self.logger.error(f"❌ 상담 데이터 입력 실패 ({row_data['학번']}): {str(e)}")
            return False

    def _input_basic_info(self, iframe, row_data):
        """기본 정보 입력 (날짜, 시간, 분야, 구분)"""
        
        # 상담일자
        if '상담일자' in row_data and pd.notna(row_data['상담일자']):
            self.logger.info("🟢 상담일자 입력")
            iframe.locator("#Pdate").fill(str(row_data['상담일자']).strip())
        
        # 상담시간
        if '상담시간_시' in row_data and pd.notna(row_data['상담시간_시']):
            self.logger.info("🟢 상담시간 입력")
            iframe.locator("input[name='Hour']").fill(str(int(row_data['상담시간_시'])))
        
        if '상담시간_분' in row_data and pd.notna(row_data['상담시간_분']):
            iframe.locator("input[name='Min']").fill(str(int(row_data['상담시간_분'])))
        
        # 상담분야 - label로 선택 (Node.js와 동일)
        if '상담분야' in row_data and pd.notna(row_data['상담분야']):
            self.logger.info("🟢 상담분야 선택")
            field = str(row_data['상담분야']).strip()
            iframe.locator("#Cntype").select_option(label=field)
        
        # 상담구분 (개인/집단)
        if '상담구분' in row_data and pd.notna(row_data['상담구분']):
            self.logger.info("🟢 상담구분 선택")
            is_group = str(row_data['상담구분']).strip() == '집단상담'
            value = '2' if is_group else '1'
            iframe.locator(f"label:has(input[name='CnPer'][value='{value}'])").click()

    def _select_counseling_type(self, iframe, counseling_type):
        """상담구분 선택 (개인상담/집단상담)"""
        counseling_category = str(counseling_type).strip()
        
        try:
            if counseling_category in ['개인상담', '개인', '1']:
                value = '1'
                type_text = "개인상담"
            elif counseling_category in ['집단상담', '집단', '2']:
                value = '2'
                type_text = "집단상담"
            else:
                self.logger.warning(f"알 수 없는 상담구분 값: {counseling_category}")
                return
            
            # 라벨 클릭으로 선택
            label_selector = f"label:has(input[name='CnPer'][value='{value}'])"
            iframe.locator(label_selector).click()
            self.logger.info(f"{type_text} 선택됨")
            time.sleep(0.3)
                
        except Exception as e:
            self.logger.error(f"상담구분 선택 중 오류: {str(e)}")

    def _input_counseling_content(self, iframe, row_data):
        """상담 내용 입력"""
        self.logger.info("🟢 제목/내용 입력")
        
        # 제목
        if '제목' in row_data and pd.notna(row_data['제목']):
            iframe.locator("#Title").fill(str(row_data['제목']).strip())
        
        # 상담내용
        if '상담내용' in row_data and pd.notna(row_data['상담내용']):
            iframe.locator("#Content").fill(str(row_data['상담내용']).strip())

    def _input_student_status(self, iframe, row_data):
        """학생상태 입력"""
        try:
            for csv_field, html_name in self.STATUS_FIELDS.items():
                if csv_field in row_data and pd.notna(row_data[csv_field]):
                    status_value = str(row_data[csv_field]).strip()
                    html_value = self.STATUS_MAPPING.get(status_value, '1')
                    
                    # 라벨 클릭으로 선택
                    label_selector = f"label:has(input[name='{html_name}'][value='{html_value}'])"
                    iframe.locator(label_selector).click()
                    self.logger.info(f"{csv_field}: {status_value} 선택됨")
                    time.sleep(0.1)
                        
        except Exception as e:
            self.logger.error(f"학생상태 입력 중 오류: {str(e)}")

    def _input_referral_status(self, iframe, row_data):
        """전문상담의뢰 입력"""
        if '전문상담의뢰' not in row_data or not pd.notna(row_data['전문상담의뢰']):
            return
            
        try:
            referral_value = str(row_data['전문상담의뢰']).strip().upper()
            value = '1' if referral_value in ['Y', '예', '1'] else '2'
            text = '예' if value == '1' else '아니오'
            
            # 라벨 클릭으로 선택
            label_selector = f"label:has(input[name='CounReq'][value='{value}'])"
            iframe.locator(label_selector).click()
            self.logger.info(f"전문상담의뢰: {text} 선택됨")
            time.sleep(0.1)
                
        except Exception as e:
            self.logger.error(f"전문상담의뢰 입력 중 오류: {str(e)}")

    def _set_privacy_option(self, iframe, row_data):
        """비공개설정"""
        if ('비공개설정' in row_data and 
            str(row_data['비공개설정']).strip().upper() == 'Y'):
            try:
                iframe.locator("input[type='checkbox']").first.click()
                time.sleep(0.3)
            except Exception as e:
                self.logger.warning(f"비공개설정 실패: {str(e)}")
                        
    def save_counseling_data(self, row_data):
        """상담 데이터 저장"""
        try:
            iframe = self.page.frame_locator("iframe")
            
            self.logger.info("🟡 저장 버튼 클릭")
            
            # Dialog 자동 처리
            def handle_dialog(dialog):
                self.logger.info(f"🟠 dialog: {dialog.message}")
                dialog.accept()
            
            self.page.once("dialog", handle_dialog)
            
            # 저장 버튼 클릭
            iframe.locator("#CounselInputBtn").click()
            
            # 2.5초 대기
            time.sleep(2.5)
            
            # 폼이 닫혔는지 확인
            try:
                pdate_visible = iframe.locator("#Pdate").is_visible()
            except:
                pdate_visible = False
            
            if pdate_visible:
                raise Exception("저장 후에도 입력 폼이 닫히지 않음 (저장 실패 가능)")
            
            self.logger.info("✅ 저장 후 입력 폼 닫힘")
            
            # 제목으로 저장 확인
            title = str(row_data.get('제목', '')).strip() if '제목' in row_data else ''
            
            if title:
                # body 대기
                body = iframe.locator("body")
                body.wait_for(state="visible")
                
                # 제목이 나타나는지 확인
                try:
                    iframe.get_by_text(title, exact=False).first.wait_for(state="visible", timeout=10000)
                    self.logger.info("✅ 상담 목록에서 제목 확인됨 (저장 확정)")
                except:
                    self.logger.warning("⚠️ 저장은 된 것 같지만, 목록에서 제목 확인 실패 (UI/탭/필터 영향 가능)")
            
            return True
                
        except Exception as e:
            self.logger.error(f"❌ 상담 데이터 저장 실패: {str(e)}")
            
            # 스크린샷
            try:
                timestamp = int(time.time() * 1000)
                student_id = row_data.get('학번', 'unknown')
                self.page.screenshot(path=f"fail_{student_id}_{timestamp}.png", full_page=True)
            except:
                pass
            
            return False
    
    def close_profile(self):
        """프로필 창 닫기"""
        try:
            self.page.keyboard.press("Escape")
            time.sleep(1)
            return True
        except Exception as e:
            self.logger.warning(f"프로필 창 닫기 실패: {str(e)}")
            return True
    
    def process_all_students(self):
        """모든 학생 상담 데이터 처리"""
        if not hasattr(self, 'df'):
            self.logger.error("CSV 데이터가 로드되지 않았습니다.")
            return False
        
        total_students = len(self.df)
        success_count = 0
        
        print(f"\n📋 2단계: 상담 데이터 자동 입력 시작")
        print(f"🎯 총 {total_students}명의 학생 상담 데이터를 처리합니다.\n")
        
        for index, row in self.df.iterrows():
            try:
                student_id = str(row['학번']).strip()
                student_name = str(row['이름']).strip()
                
                print(f"📝 처리 중: [{index+1}/{total_students}] {student_name}({student_id})")
                
                result = {
                    'index': index + 1,
                    'student_id': student_id,
                    'student_name': student_name,
                    'status': 'FAILED',
                    'error_message': '',
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                
                # 순차 처리
                steps = [
                    (self.search_student, student_id, '학생 검색 실패'),
                    (self.open_student_profile, student_id, '프로필 열기 실패'),
                    (self.input_counseling_data, row, '상담 데이터 입력 실패')
                ]
                
                success = True
                for step_func, param, error_msg in steps:
                    if not step_func(param):
                        result['error_message'] = error_msg
                        success = False
                        break
                
                # 여기가 핵심! row를 전달
                if success and self.save_counseling_data(row):  # ← row 추가!
                    result['status'] = 'SUCCESS'
                    success_count += 1
                    print(f"   ✅ 성공: 상담 데이터 입력 및 저장 완료")
                elif success:
                    result['error_message'] = '상담 데이터 저장 실패'
                    print(f"   ⚠️ 입력 완료, 저장 확인 필요")
                else:
                    print(f"   ❌ 실패: {result['error_message']}")
                
                self.close_profile()
                self.results.append(result)
                time.sleep(2)
                
            except Exception as e:
                error_msg = f"예상치 못한 오류: {str(e)}"
                print(f"   ❌ 오류: {error_msg}")
                
                result = {
                    'index': index + 1,
                    'student_id': student_id if 'student_id' in locals() else 'Unknown',
                    'student_name': student_name if 'student_name' in locals() else 'Unknown',
                    'status': 'ERROR',
                    'error_message': error_msg,
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                self.results.append(result)
        
        print(f"\n🎉 전체 처리 완료!")
        print(f"✅ 성공: {success_count}건 / 전체: {total_students}건")
        return True
    
    def generate_report(self):
        """처리 결과 리포트 생성"""
        try:
            print(f"\n📋 3단계: 처리 결과 리포팅")
            
            total = len(self.results)
            success = len([r for r in self.results if r['status'] == 'SUCCESS'])
            failed = len([r for r in self.results if r['status'] == 'FAILED'])
            error = len([r for r in self.results if r['status'] == 'ERROR'])
            
            print("\n" + "="*70)
            print("📊 상담 데이터 자동 입력 결과 리포트")
            print("="*70)
            print(f"📈 전체 처리 건수: {total}건")
            print(f"✅ 성공: {success}건 ({success/total*100:.1f}%)")
            print(f"❌ 실패: {failed}건 ({failed/total*100:.1f}%)")
            print(f"⚠️ 오류: {error}건 ({error/total*100:.1f}%)")
            print("-"*70)
            
            # 실패/오류 항목 출력
            if failed > 0 or error > 0:
                print("📋 실패/오류 상세 내역:")
                for i, result in enumerate(self.results):
                    if result['status'] != 'SUCCESS':
                        status_icon = "❌" if result['status'] == 'FAILED' else "⚠️"
                        print(f"  {i+1:2d}. {status_icon} {result['student_name']}({result['student_id']}): {result['error_message']}")
                print("-"*70)
            
            # CSV 리포트 생성
            report_df = pd.DataFrame(self.results)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            report_filename = f"cando_auto_report_{timestamp}.csv"
            report_df.to_csv(report_filename, index=False, encoding='utf-8-sig')
            
            print(f"📁 상세 리포트 파일 생성: {report_filename}")
            print("="*70)
            
            return True
            
        except Exception as e:
            self.logger.error(f"리포트 생성 실패: {str(e)}")
            return False
    
    def cleanup(self):
        """리소스 정리"""
        try:
            if hasattr(self, 'page') and self.page:
                self.page.close()
            if hasattr(self, 'browser') and self.browser:
                self.browser.close()
            if hasattr(self, 'playwright') and self.playwright:
                self.playwright.stop()
            self.logger.info("브라우저 리소스 정리 완료")
        except Exception as e:
            self.logger.error(f"리소스 정리 중 오류: {str(e)}")
    
    def run(self):
        """메인 실행 함수"""
        try:
            steps = [
                (self.load_csv_data, "CSV 데이터 로드"),
                (self.setup_browser, "브라우저 설정"),
                (self.wait_for_login, "로그인 대기"),
                (self.process_all_students, "학생 데이터 처리")
            ]
            
            for step_func, step_name in steps:
                if not step_func():
                    self.logger.error(f"{step_name} 실패")
                    return False
            
            self.generate_report()
            
            print("\n🎉 프로그램이 성공적으로 완료되었습니다!")
            print("👋 브라우저 창을 닫으려면 엔터키를 눌러주세요...")
            input()
            
            return True
            
        except KeyboardInterrupt:
            print("\n\n⚠️ 사용자에 의해 프로그램이 중단되었습니다.")
            return False
        except Exception as e:
            self.logger.error(f"프로그램 실행 중 오류: {str(e)}")
            return False
        finally:
            self.cleanup()

def main():
    """메인 함수"""
    print("🎓 호서대학교 Cando 시스템 자동 상담 입력 프로그램")
    print("="*60)
    
    while True:
        csv_path = input("📁 CSV 파일 경로를 입력하세요 (예: data.csv): ").strip().strip('"')
        
        if os.path.exists(csv_path):
            break
        else:
            print("❌ 파일이 존재하지 않습니다. 다시 입력해주세요.")
    
    auto_counseling = CandoAutoCounseling(csv_path)
    success = auto_counseling.run()
    
    if success:
        print("✅ 프로그램이 정상적으로 완료되었습니다.")
    else:
        print("❌ 프로그램 실행 중 오류가 발생했습니다.")
    
    return success

if __name__ == "__main__":
    main()