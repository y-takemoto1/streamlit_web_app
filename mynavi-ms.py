import time
import openpyxl
import requests
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
import streamlit as st
import sys
import os


# スクレイピングしたいURL（例: 福岡の求人）
url = 'https://mynavi-ms.jp/search/fukuoka/area-all'
headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'
        }

# エクセルを開く
wb = openpyxl.Workbook()
ws = wb.active

# 保存するフォルダのパス
folder_path = 'C:/aggregate'
# 保存するファイル名
file_name = 'mynavi-ms.xlsx'
# 完全なファイルパス
file_path = os.path.join(folder_path, file_name)

st.title('マイナビシニアの求人検索')

# 初期の値を表示するための空の場所を作成
text_placeholder = st.empty()

# 初期化
if 'stop' not in st.session_state:
    st.session_state.stop = False
if 'processing' not in st.session_state:
    st.session_state.processing = False


if st.button('開始', disabled=st.session_state.processing):
    st.session_state.stop = False
    # 中断ボタン

    if url:
        # ステータスコードをチェック
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            c = 1  # 行数
            j = 1  # ページ数
            while c <= 10:
                print(f'Processing page {j}...')
                soup = BeautifulSoup(response.text, 'html.parser')
                job_cards = soup.find_all('div', class_='job-summary__header')
                table_cards = soup.find_all('table', class_='job-summary-table')

                if not job_cards:
                    print('求人が見つかりませんでした。')
                    break
                for job, job_table in zip(job_cards, table_cards):
                    print('testcount:', c)

                    if st.session_state.stop:
                        st.session_state.running = False
                        st.warning("処理が中断されました。")
                        break


                    # プレースホルダーに新しいテキストを表示
                    text_placeholder.text(f'読込数：{c}')

                    # 求人タイトルを取得
                    title = job.find('a', target='_blank').text.strip()
                    print('求人タイトル:', title)
                    
                    # 会社名を取得
                    company = job.find('p', class_='job-summary__name').text.strip()
                    print('会社名:', company)

                    # テーブルデータを取得（例として最初のテーブルを取得）
                    rows = job_table.find_all('tr')

                    # テーブルデータをリストに格納
                    table_data = []
                    for row in rows:
                        cols = row.find_all('td')
                        cols = [col.get_text(strip=True) for col in cols]
                        if cols:  # 空でない場合のみ追加
                            table_data.append(cols)

                    # 1つのセルにデータを書き込むための文字列を作成
                    cell_data = '\n'.join(['\t'.join(row) for row in table_data])


                    # 詳細ページのリンクを取得
                    job_link = 'https://mynavi-ms.jp/' + job.find('a')['href']

                    # 待機
                    time.sleep(10)
                    # 詳細ページにリクエストを送信
                    job_response = requests.get(job_link)
                    job_soup = BeautifulSoup(job_response.text, 'html.parser')

                    # 電話番号を正規表現で抽出（例: 012-345-6789）
                    phone_numbers = re.findall(r'\d{2,4}-\d{2,4}-\d{2,4}', job_soup.text)
                    phone_numbers2 = [number for number in phone_numbers if not number.startswith("0120")]
                    if phone_numbers2:
                        print('電話番号:', phone_numbers)
                        # Excelファイルに書き込mi 
                        ws.cell(row=c, column=1, value=title)
                        ws.cell(row=c, column=2, value=company)
                        ws.cell(row=c, column=3, value=', '.join(phone_numbers) if phone_numbers else 'なし')
                        ws.cell(row=c, column=4, value=cell_data)
                        c += 1
                    else:
                        print('電話番号は見つかりませんでした。')
                    
                    print('-' * 40)
                # 次のページを探す
                next_page = soup.find('a', class_='search-pagenation__next')
                if next_page and 'href' in next_page.attrs:
                    next_url = 'https://mynavi-ms.jp' + next_page['href']
                    response = requests.get(next_url, headers=headers)
                    time.sleep(5)  # サーバーへの負荷を避けるためにスリープ
                    j += 1
                else:
                    print('pageEnd')
                    break  # 次のページがない場合は終了
            else:
                st.text('完了しました。')
                st.text("C:/aggregate/に保存されました。")
        else:
            print('リクエストが失敗しました。ステータスコード:', response.status_code)

# フォルダの有無確認
if not os.path.exists(folder_path):
    # フォルダ生成
    os.mkdir(folder_path)

# エクセルファイルを保存
wb.save(file_path)
wb.close()
st.session_state.processing = False

if st.button("中断", disabled=st.session_state.processing):
    st.session_state.stop = True