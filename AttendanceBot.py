
import time
import os
import configparser
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select

class AttendanceBot:
  def __init__(self):
    """___init___コンストラクタ（初期化）"""
    try:
      # Chrome WebDriverを起動
      self.driver = webdriver.Chrome()
      self.load_env()
      self.load_config()
    except Exception as e:
      print(f'Initialization Error: {e}')
      self.driver.quit()

  def load_env(self):
    """環境変数ファイル(.env)読み込み"""
    try:
      load_dotenv()
      self.user_cd = os.getenv("USER_CD")
      self.user_id = os.getenv("USER_ID")
      self.user_password = os.getenv("USER_PASSWORD")
    except Exception as e:
      print(f"Environment Load Error: {e}")

  def load_config(self):
    """設定ファイル(.ini)読み込む"""
    try:
      config = configparser.ConfigParser()
      config.read("config.ini")
      self.holiday = config["attendance"]["holiday"]
      self.work_remotely = config["attendance"]["work_remotely"]
      self.parse_config()
    except Exception as e:
      print(f"Config Load Error: {e}")

  def parse_config(self):
    """読み込んだ設定ファイルをリストに変換する"""
    try:
      if self.holiday:
        self.holiday_list = self.holiday.split(',')
      if self.work_remotely:
        self.work_remotely_list = self.work_remotely.split(',')
    except Exception as e:
      print(f'parse_config Error: {e}')

  def login(self):
    """勤怠システムにログインする"""
    try:
      # 勤怠管理システムのURLを指定
      self.driver.get('https://vsn.digisheet.com/staffLogin')
      self.driver.find_element(By.NAME, "HC").send_keys(self.user_cd)
      self.driver.find_element(By.NAME, "UI").send_keys(self.user_id)
      self.driver.find_element(By.NAME, "Pw").send_keys(self.user_password)
      self.driver.find_element(By.NAME, "loginButton").click()
    except Exception as e:
      print(f"Login Error: {e}")

  def switch_to_frame(self, frame_name):
    """フレームを切り替える"""
    try:
      # デフォルトのframeに切り替える
      self.driver.switch_to.default_content()
      # フレームを切り替える
      self.driver.switch_to.frame(frame_name)
    except Exception as e:
      print(f"switch_to_frame Error ({frame_name}): {e}")

  ### 勤怠報告ページ移動
  def navigate_to_attendance(self):
    try:
      self.switch_to_frame("menu")
      self.driver.find_element(By.LINK_TEXT, "勤務報告").click()
      self.switch_to_frame("main")
    except Exception as e:
      print(f"navigate_to_attendance Error: {e}")

  ### 勤怠表のテーブル要素を取得
  def get_attendance_table(self):
    try:
      return WebDriverWait(self.driver, 10).until(
          EC.presence_of_element_located((By.XPATH, "/html/body/form/table/tbody/tr[7]/td/table"))
      )
    except Exception as e:
      print(f"get_attendance_table Error: {e}")
      return None

  def input_attendance(self):
    """勤怠情報を入力する"""
    try:
      # 勤怠テーブル要素を取得
      table_element = self.get_attendance_table()
      if not table_element:
        return

      # 勤怠テーブルの全ての行要素を取得
      tr_elements = table_element.find_elements(By.XPATH, ".//tr")
      # 勤怠テーブルの各行を繰り返す
      for i in range(len(tr_elements)):
        tr_element = tr_elements[i]
        # 背景色を取得
        tr_bgcolor = tr_element.get_attribute("bgcolor")
        # 行の背景色が「white」か確認（平日かを判定）
        if tr_bgcolor and tr_bgcolor == "white":
          # 勤怠テーブルの各列を繰り返す
          for td_element in tr_element.find_elements(By.XPATH, ".//td"):
            # 背景色を取得
            td_bgcolor = td_element.get_attribute("bgcolor")
            # セルの背景色が「#0000FF」（日付の部分か判定）
            if td_bgcolor and td_bgcolor == "#0000FF":
              input_date = td_element.text
              print(f'{input_date}日') # 日付ログ
              # 日付セルをクリック
              td_element.click()

              time.sleep(1) # 画面描画のため待機
              # 勤怠登録
              self.regist_attendance_form(input_date)
              time.sleep(1) # 画面描画のため待機

              # 画面遷移後に勤怠テーブル要素を再取得
              table_element = self.get_attendance_table()
              if not table_element:
                return

              # 画面遷移後に勤怠テーブルの全ての行要素を取得
              tr_elements = table_element.find_elements(By.XPATH, ".//tr")
              break
    except Exception as e:
      print(f'input_attendance Error: {e}')

  def regist_attendance_form(self, day):
    """勤怠フォームに登録する"""
    try:
      if day in self.holiday_list:
        # 届出選択
        report_select = Select(self.driver.find_element(By.NAME, "AttendSecSelect"))
        # 「年次有給休暇（有給）」を選択
        report_select.select_by_value("12")
        time.sleep(1)
      else:
        if day in self.work_remotely_list:
          # その他業務
          other_work_select = Select(self.driver.find_element(By.NAME, "ContentSelect"))
          # 「在宅　所定時間以上」を選択
          other_work_select.select_by_value("0000000600")
          time.sleep(1)

        # 勤務シフト選択
        shift_select = Select(self.driver.find_element(By.NAME, "AttendSelect"))
        # 「B勤務」を選択
        shift_select.select_by_value("B0")
        time.sleep(1)

      # 登録ボタンクリック
      self.driver.find_element(By.XPATH, "//input[@value='登　録']").click()
    except Exception as e:
      print(f"regist_attendance Error: {e}")

  def run(self):
    """勤怠入力の処理を実行する"""
    try:
      self.login()
      self.navigate_to_attendance()
      self.input_attendance()
    except Exception as e:
      print(f'Exception Error: {e}')
    finally:
      self.driver.quit()


### Pythonスクリプトが直接実行されたときに動作するエントリーポイント
### スクリプトが実行されると「__name__」は「__main__」になる
if __name__ == "__main__":
  # AttendanceBotクラスのインスタンスを作成
  bot = AttendanceBot()
  # runメソッドを実行
  bot.run()