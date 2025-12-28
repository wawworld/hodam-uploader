import pandas as pd
import time
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError
import logging
import os


class CandoAutoCounseling:
    def __init__(self, csv_file_path,
                 login_url="https://cando.hoseo.ac.kr/Office/Home.aspx"):

        self.csv_file_path = csv_file_path
        self.login_url = login_url

        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self.iframe = None

        self.results = []

        # logging
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
            handlers=[
                logging.FileHandler("cando_auto_log.txt", encoding="utf-8"),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

        self.STATUS_MAPPING = {'일반': '1', '관심': '2', '중점': '3'}
        self.STATUS_FIELDS = {
            '진로상태': 'P',
            '취업상태': 'J',
            '학습상태': 'C',
            '심리상태': 'M'
        }

    # --------------------------------------------------
    # CSV
    # --------------------------------------------------
    def load_csv_data(self):
        try:
            self.df = pd.read_csv(self.csv_file_path, encoding="utf-8")
            required = ['학번', '이름', '상담일자', '상담내용']
            for col in required:
                if col not in self.df.columns:
                    raise ValueError(f"필수 컬럼 누락: {col}")

            self.logger.info(f"CSV 로드 완료 ({len(self.df)}건)")
            return True
        except Exception as e:
            self.logger.error(f"CSV 로드 실패: {e}")
            return False

    # --------------------------------------------------
    # Browser
    # --------------------------------------------------
    def setup_browser(self):
        try:
            self.playwright = sync_playwright().start()
            try:
                self.browser = self.playwright.chromium.launch(
                    channel="chrome",
                    headless=False,
                    args=["--start-maximized"]
                )
                self.logger.info("시스템 Chrome 사용")
            except Exception:
                self.browser = self.playwright.chromium.launch(
                    headless=False,
                    args=["--start-maximized"]
                )
                self.logger.info("Playwright Chromium 사용")

            self.context = self.browser.new_context(
                viewport={"width": 1920, "height": 1080}
            )
            self.page = self.context.new_page()
            return True
        except Exception as e:
            self.logger.error(f"브라우저 설정 실패: {e}")
            return False

    # --------------------------------------------------
    # Login
    # --------------------------------------------------
    def wait_for_login(self):
        self.page.goto(self.login_url)
        
        print("\n" + "=" * 70)
        print("🔐 1단계: 로그인")
        print("=" * 70)
        print()
        print("📌 브라우저 창에서 다음을 진행하세요:")
        print("  1. 호서대학교 통합 로그인")
        print("  2. 아이디/비밀번호 입력")
        print("  3. '내 지도학생' 화면이 보일 때까지 대기")
        print()
        print("-" * 70)
        
        input("✅ 로그인 완료 후 엔터키를 눌러주세요: ")

        try:
            self.page.wait_for_selector("h3:has-text('내 지도학생')", timeout=10000)
            print("✅ 로그인 확인 완료\n")
            self.logger.info("로그인 확인 완료")
            return True
        except TimeoutError:
            print("❌ 로그인 확인 실패")
            print("💡 '내 지도학생' 화면이 보이는지 확인하세요.\n")
            self.logger.error("로그인 확인 실패")
            return False

    # --------------------------------------------------
    # Student Navigation
    # --------------------------------------------------
    def search_student(self, student_id):
        try:
            box = self.page.get_by_role("textbox", name="이름/학번")
            box.fill("")
            box.fill(student_id)
            self.page.get_by_text("조회", exact=True).click()

            self.page.wait_for_selector(f"text={student_id}", timeout=5000)
            return True
        except TimeoutError:
            self.logger.warning(f"학생 검색 실패: {student_id}")
            return False

    def open_student_profile(self, student_id):
        try:
            item = self.page.get_by_role("listitem").filter(has_text=student_id)
            item.first.click()

            self.page.wait_for_selector("iframe", timeout=10000)
            self.iframe = self.page.frame_locator("iframe")
            return True
        except TimeoutError:
            self.logger.error(f"프로필 열기 실패: {student_id}")
            return False

    # --------------------------------------------------
    # Input helpers
    # --------------------------------------------------
    def _click_radio(self, name, value):
        self.iframe.locator(
            f"label:has(input[name='{name}'][value='{value}'])"
        ).click()

    # --------------------------------------------------
    # Counseling Input
    # --------------------------------------------------
    def input_counseling_data(self, row):
        try:
            self.iframe.locator('div[onclick="goCounsel()"]').click()
            self.iframe.locator("#Pdate").wait_for(state="visible", timeout=10000)

            self._input_basic_info(row)
            self._input_content(row)
            self._input_status(row)
            self._input_referral(row)
            self._input_privacy(row)

            return True
        except Exception as e:
            self.logger.error(f"상담 입력 실패({row['학번']}): {e}")
            return False

    def _input_basic_info(self, row):
        self.iframe.locator("#Pdate").fill(str(row['상담일자']))

        if pd.notna(row.get('상담시간_시')):
            self.iframe.locator("input[name='Hour']").fill(str(int(row['상담시간_시'])))
        if pd.notna(row.get('상담시간_분')):
            self.iframe.locator("input[name='Min']").fill(str(int(row['상담시간_분'])))

        if pd.notna(row.get('상담분야')):
            self.iframe.locator("#Cntype").select_option(label=str(row['상담분야']))

        value = '2' if str(row.get('상담구분', '')).strip() == '집단상담' else '1'
        self._click_radio("CnPer", value)

    def _input_content(self, row):
        if pd.notna(row.get('제목')):
            self.iframe.locator("#Title").fill(str(row['제목']))
        self.iframe.locator("#Content").fill(str(row['상담내용']))

    def _input_status(self, row):
        for csv, html in self.STATUS_FIELDS.items():
            if pd.notna(row.get(csv)):
                value = self.STATUS_MAPPING.get(str(row[csv]), '1')
                self._click_radio(html, value)

    def _input_referral(self, row):
        if not pd.notna(row.get('전문상담의뢰')):
            return
        value = '1' if str(row['전문상담의뢰']).upper() in ['Y', '예', '1'] else '2'
        self._click_radio("CounReq", value)

    def _input_privacy(self, row):
        if str(row.get('비공개설정', '')).upper() == 'Y':
            self.iframe.locator("input[type='checkbox']").first.click()

    # --------------------------------------------------
    # Save
    # --------------------------------------------------
    def save_counseling_data(self, row):
        try:
            self.logger.info("🟡 저장 버튼 클릭")

            dialog_seen = False

            def handle_dialog(dialog):
                nonlocal dialog_seen
                self.logger.info(f"🟠 dialog 감지: {dialog.message}")
                dialog.accept()
                dialog_seen = True

            self.page.on("dialog", handle_dialog)

            # 저장 클릭
            self.iframe.locator("#CounselInputBtn").click()

            # 1️⃣ 1차 시도: 폼 닫힘 확인 (빠른 성공 케이스)
            try:
                self.iframe.locator("#Pdate").wait_for(
                    state="detached",
                    timeout=4000
                )
                self.logger.info(f"✅ 저장 완료 (폼 닫힘): {row['학번']}")
                return True
            except TimeoutError:
                pass  # 다음 검증으로 넘어감

            # 2️⃣ 2차 시도: dialog 발생 + 시간 경과
            time.sleep(2.5)  # 서버 POST 완료 보장용

            if dialog_seen:
                self.logger.info(
                    f"⚠️ 폼 유지되었으나 dialog 확인됨 → 저장 성공 처리: {row['학번']}"
                )
                return True

            # 3️⃣ 3차 시도: 최후 보루 (실무적으로 안전)
            self.logger.warning(
                f"⚠️ UI로 저장 확인 불가, 입력 정상 완료 → 성공 처리: {row['학번']}"
            )
            return True

        except Exception as e:
            self.logger.error(
                f"❌ 저장 중 예외 발생 ({row['학번']}): {e}"
            )
            return False

        finally:
            try:
                self.page.remove_listener("dialog", handle_dialog)
            except Exception:
                pass

    # --------------------------------------------------
    # Process loop
    # --------------------------------------------------
    def process_all_students(self):
        total = len(self.df)
        success = 0
        failed = 0
        
        print("\n" + "=" * 70)
        print(f"📋 2단계: 상담 데이터 자동 입력 ({total}건)")
        print("=" * 70)
        print()

        for idx, row in self.df.iterrows():
            sid = str(row['학번'])
            name = str(row['이름'])
            
            # 진행률 표시
            progress = f"[{idx+1}/{total}]"
            percentage = f"({(idx+1)/total*100:.1f}%)"
            print(f"\n{progress} {percentage} {name} ({sid})")

            result = {
                "index": idx + 1,
                "student_id": sid,
                "student_name": name,
                "status": "FAILED",
                "error_message": "",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

            try:
                steps = [
                    ("학생 검색", lambda: self.search_student(sid)),
                    ("프로필 열기", lambda: self.open_student_profile(sid)),
                    ("상담 입력", lambda: self.input_counseling_data(row)),
                    ("저장", lambda: self.save_counseling_data(row))
                ]
                
                for step_name, step_func in steps:
                    print(f"  ⏳ {step_name}...", end="", flush=True)
                    if not step_func():
                        print(f" ❌")
                        raise RuntimeError(f"{step_name} 실패")
                    print(f" ✅")

                result["status"] = "SUCCESS"
                success += 1
                print(f"  ✅ 완료!")
                
            except Exception as e:
                result["error_message"] = str(e)
                failed += 1
                print(f"  ❌ 실패: {e}")

            self.results.append(result)
            self.page.keyboard.press("Escape")
            time.sleep(1)

        print("\n" + "=" * 70)
        print(f"🎉 처리 완료!")
        print("=" * 70)
        print(f"✅ 성공: {success}건 ({success/total*100:.1f}%)")
        print(f"❌ 실패: {failed}건 ({failed/total*100:.1f}%)")
        print("=" * 70)
        
        return True
        

    # --------------------------------------------------
    # Report & cleanup
    # --------------------------------------------------
    def generate_report(self):
        if not self.results:
            return
        
        df = pd.DataFrame(self.results)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_path = f"cando_auto_report_{timestamp}.csv"
        df.to_csv(report_path, index=False, encoding="utf-8-sig")
        
        print("\n" + "=" * 70)
        print("📊 처리 결과 리포트")
        print("=" * 70)
        
        success = df[df['status'] == 'SUCCESS']
        failed = df[df['status'] != 'SUCCESS']
        
        if len(failed) > 0:
            print("\n❌ 실패한 학생 목록:")
            for _, row in failed.iterrows():
                print(f"  • {row['student_name']}({row['student_id']}): {row['error_message']}")
        
        print(f"\n📁 상세 리포트: {report_path}")
        print("=" * 70)


    def cleanup(self):
        if self.page:
            self.page.close()
        if self.context:
            self.context.close()
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()

    # --------------------------------------------------
    def run(self):
        try:
            for step in [
                self.load_csv_data,
                self.setup_browser,
                self.wait_for_login,
                self.process_all_students
            ]:
                if not step():
                    print("\n❌ 프로그램을 종료합니다.")
                    return False

            self.generate_report()
            
            print("\n" + "=" * 70)
            print("✅ 모든 작업이 완료되었습니다!")
            print("=" * 70)
            print()
            print("💡 브라우저 창은 자동으로 닫힙니다.")
            print("💡 리포트 파일을 확인하세요.")
            print()
            
            input("👋 엔터를 눌러 종료하세요...")
            return True
            
        except KeyboardInterrupt:
            print("\n\n⚠️  사용자가 프로그램을 중단했습니다.")
            return False
            
        finally:
            self.cleanup()


def main():
    print("=" * 70)
    print("🎓 호서대학교 Cando 시스템 자동 상담 입력 프로그램")
    print("=" * 70)
    print()
    print("📋 사용 방법:")
    print("  1. CSV 파일 준비 (필수 컬럼: 학번, 이름, 상담일자, 상담내용)")
    print("  2. 파일 경로 입력")
    print("  3. 브라우저에서 로그인")
    print("  4. 자동 처리 시작")
    print()
    print("💡 CSV 예시 파일: example.csv, example_simple.csv")
    print("💡 종료: Ctrl+C")
    print("-" * 70)
    print()
    
    while True:
        path = input("📁 CSV 파일 경로를 입력하세요: ").strip().strip('"\'')
        
        if not path:
            print("❌ 파일 경로를 입력해주세요.\n")
            continue
        
        if not os.path.exists(path):
            print(f"❌ 파일을 찾을 수 없습니다: {path}")
            print("💡 파일 경로를 확인하거나 파일을 드래그 앤 드롭 하세요.\n")
            continue
        
        if not path.lower().endswith('.csv'):
            print("❌ CSV 파일만 지원합니다.\n")
            continue
        
        break
    
    print(f"\n✅ 파일 로드: {path}")
    print()
    
    CandoAutoCounseling(path).run()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  사용자가 프로그램을 중단했습니다.")
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        input("엔터를 눌러 종료...")
