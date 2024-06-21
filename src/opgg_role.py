import csv
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import re

# Chrome WebDriverの設定
chrome_options = Options()
chrome_options.add_argument("--disable-popup-blocking")  # ポップアップブロックを無効化
service = Service("path/to/chromedriver.exe")  # パスを適切なものに変更
driver = uc.Chrome(service=service, options=chrome_options)

# ウェブページを開く
driver.get("https://www.op.gg/leaderboards/tier?tier=iron&page=1")

# ファイル名を作成するためのURLからの文字列取得
url = driver.current_url
start_idx = url.find("tier=") + len("tier=")
end_idx = url.find("&", start_idx)
file_name = url[start_idx:end_idx]

# 取得する人数を指定
num_of_players = 50

# ランキング情報を格納するリスト
rankings = []

for i in range(1, num_of_players + 1):
    # 指定された要素から情報を取得
    xpath = f'/html/body/div[1]/div[6]/div[3]/table/tbody/tr[{i}]'
    row = driver.find_element(By.XPATH, xpath)

    # 各セルのテキストを取得
    cells = row.find_elements(By.TAG_NAME, "td")

    # サモナー名とURLを取得
    summoner_cell = cells[1]
    summoner_link = summoner_cell.find_element(By.CLASS_NAME, "summoner-link")
    summoner_name = summoner_link.find_element(By.CLASS_NAME, "css-ao94tw").text
    summoner_tag = summoner_link.find_element(By.CLASS_NAME, "css-1mbuqon").text
    summoner_url = summoner_link.get_attribute("href")

    # サモナー名に # 以降の文字列を追加
    summoner_name = f"{summoner_name} {summoner_tag}"

    # ランキング情報をリストに追加
    ranking_info = []
    for cell in cells:
        ranking_info.append(cell.text)

    # サモナーURLを追加
    ranking_info[4] = summoner_url

    # サモナー名を追加
    ranking_info[1] = summoner_name

    # ロール情報を追加
    role_cell = cells[5]
    role_links = role_cell.find_elements(By.TAG_NAME, "a")
    roles = [role_link.get_attribute("href") for role_link in role_links]
    role_info = ", ".join(roles)
    ranking_info.insert(7, role_info)

    # ランキング情報をリストに追加
    rankings.append(ranking_info)

# 各サモナーのウェブページを新しいタブで開いて情報を取得
for ranking in rankings:
    summoner_url = ranking[4]
    summoner_name = ranking[1]
    
    # サモナーのウェブページを新しいタブで開く
    driver.execute_script(f"window.open('{summoner_url}', '_blank');")
    time.sleep(4)  # タブが完全にロードされるのを待つ

    # 新しいタブに切り替える
    driver.switch_to.window(driver.window_handles[-1])
    
    try:
        # データ更新
        button = driver.find_element(By.XPATH, "/html/body/div[1]/div[4]/div[1]/div/div[1]/div[2]/div[5]/button[1]")
        button.click()
        time.sleep(5)  # データが更新されるのを待つ
    except:
        time.sleep(0)  # データが更新されるのを待つ
        print(f"{summoner_name} データ更新ができませんでした ")
    
    try:
        # ランク (ソロ/デュオ)を選択
        rank_button = driver.find_element(By.XPATH, '//*[@id="content-container"]/div[2]/div[1]/ul/li[2]/button')
        rank_button.click()
        time.sleep(2)  # ページが読み込まれるのを待つ
    except:
        time.sleep(0)
        print(f"{summoner_name} ランク (ソロ/デュオ)を選択ができませんでした ")

    # スタッツからパーセンテージ値を取得
    roles = ["top", "jg", "mid", "adc", "sup"]
    style_values = []
    for role in roles:
        try:
            xpath = f'//*[@id="content-container"]/div[2]/div[2]/div[3]/ul/li[{roles.index(role) + 1}]/div[1]/div'
            element = driver.find_element(By.XPATH, xpath)
            style = element.get_attribute("style")
            match = re.search(r"height:\s*([\d.]+)%", style)
        except:
            print(f"{summoner_name} ピックレートが取得できませんでした ")

        if match:
            percentage = match.group(1)
            #print(f"{summoner_name} - {role}: {percentage}%")
            style_values.append(percentage)
        else:
            style_values.append("N/A")

    # タブを閉じる
    driver.close()

    # 元のタブに戻る
    driver.switch_to.window(driver.window_handles[0])

    # ランキング情報にstyle値を追加
    ranking.extend(style_values)

# キー入力があるまで待機
input("Enterを押すと終了します。")

# WebDriverを閉じる
driver.quit()

# ランキング情報を表示
for ranking in rankings:
    rank = ranking[0]
    summoner_name = ranking[1]
    tier = ranking[2]
    lp = ranking[3]
    summoner_url = ranking[4]
    level = ranking[5]
    win_rate = ranking[6]
    role = ranking[7]
    top = ranking[8]
    jg = ranking[9]
    mid = ranking[10]
    adc = ranking[11]
    sup = ranking[12]

    print("順位:", rank)
    print("サモナー名:", summoner_name)
    print("Tier:", tier)
    print("LP:", lp)
    print("サモナーURL:", summoner_url)
    print("レベル:", level)
    print("勝率:", win_rate)
    print("=INDEX($I$1:$M$1, MATCH(MAX($I2:$M2), $I2:$M2, 0))をスプレッドシートに入力してください:", role)
    print("Top:", top)
    print("Jungle:", jg)
    print("Mid:", mid)
    print("ADC:", adc)
    print("Support:", sup)
    print("-----------------------")

# CSVファイルに書き込み
csv_file_path = f"D:\Downloads\{file_name}_rankings.csv"  # 保存先のパス
with open(csv_file_path, mode='w', encoding='utf-8', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["#", "サモナー", "Tier", "LP", "サモナーURL", "レベル", "勝率", "ロール", "Top", "Jungle", "Mid", "ADC", "Support","=INDEX($I$1:$M$1, MATCH(MAX($I2:$M2), $I2:$M2, 0))をロールに入れてね"])
    writer.writerows(rankings)

print("ランキング情報が '{}' に保存されました。".format(csv_file_path))
