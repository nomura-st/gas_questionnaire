import requests
import re
from bs4 import BeautifulSoup  # BeautifulSoupクラスをインポート
import datetime
import gspread

# ★★★ ↓↓↓コマンド用↓↓↓ ★★★
import sys
gc = gspread.service_account(filename=sys.argv[1])
# ★★★ ↑↑↑コマンド用↑↑↑ ★★★

#############################
# カスタム情報
#############################
# 読み書き対象スプレッドシートID
ssID = '11HxD-a2Sq8DXPEEiVfZmNtEqkdOxJKZckPRkkT0GOJY'
# 読み書き対象スプレッドシート シート名
sheetName = '映画館'
# データ表が始まる行
rowStart = 2

# 検索対象サイト
baseURL = 'https://eiga.com'
# 対象の県
prefURL = f'{baseURL}/theater/10'
paths = {
    "ユナイテッド・シネマ 前橋": "/100101/3229/",
    "109シネマズ高崎": "/100201/3230/",
    "イオンシネマ高崎": "/100201/3231/",
}


#############################
# 関数
#############################
def select(url, selector):
    html = requests.get(url)
    # 取得したHTMLをBeautifulSoupを使ってパースします。
    soup = BeautifulSoup(html.content, "html.parser")
    # パースしたHTMLから特定の要素を抽出します。結果はリストで返ってきます。
    elem = soup.select(selector)

    return elem


def getMoviesFromEiga(url, path):
    # 1つの映画館ごとの映画情報取得
    movies = select(url + path, "main .content-container section[data-title]")
    obj = []
    print("*********************************************")
    print(f'GET INFO from {path}')

    for movie in filter(lambda m: len(m.select("h2 a")) > 0, movies):
        # 1つの映画ごとの情報を取得
        print(f'TRY {movie["data-title"]}')
        # .movie-schedule => 1つの映画のスケジュール表
        m = {
            "name": movie["data-title"],
            "schedules": list(map(lambda t: {
                "type": t.select(".movie-type")[0].get_text("/") if len(t.select(".movie-type")) > 0 else "",
                "date": list(),
            }, movie.select(".movie-schedule"))),
            "time": "0",
        }

        for td in movie.select(".movie-schedule td[data-date]"):
            dateStr = td["data-date"]
            print("  対象日付 > " + dateStr)

            if len(td.select("[data-time]")) > 0:
                # Unix時間あり(予約できる日付)
                m["schedules"][0]["date"].extend(
                    list(map(lambda d: d["data-time"],
                         td.select("[data-time]")))
                )
            else:
                # Unix時間なしのため、文字列から時間を生成
                for span in td.select("span"):
                    # span 直下の文字列が日付の場合のみ追加
                    timeStr = str(span.get_text()).strip()
                    print("***TEST***" + timeStr)
                    if re.match("^(\d+):(\d+)", timeStr):
                        timeStr = re.sub("～.*", "", timeStr).strip()
                        print("    対象時刻(追加分) > " + timeStr)
                        dt = datetime.datetime.strptime(
                            dateStr + " " + timeStr + ' +0900', '%Y%m%d %H:%M %z')
                        # TODO: schedulesはlistである必要なさそう？
                        m["schedules"][0]["date"].append(int(dt.timestamp()))

        print(
            f'  => GET {"/".join(list(map(str, m["schedules"][0]["date"])))}')

        if len(movie.select(".movie-image img")) > 0:
            m["image"] = movie.select(".movie-image img")[0]["src"]

        for i in movie.select(".movie-image .data span"):
            timeResult = re.match("(\d+)[分]*", i.get_text().strip())
            if timeResult:
                m["time"] = timeResult.group(1)

        if len(movie.select("h2 a")) > 0:
            m["link"] = baseURL + movie.select("h2 a")[0]["href"]

        obj.append(m)

    return obj


#############################
# メイン処理開始
#############################

##### 映画館ごとの上映中映画情報一覧を取得 #####
theaters = list(map(lambda name: {
    "name": name,
    "movies": getMoviesFromEiga(prefURL, paths[name]),
}, paths.keys()))


##### 取得した情報をSpreadSheetsに出力 #####
ss = gc.open_by_key(ssID)
sheet = ss.worksheet(sheetName)

# 既存を削除
ss.values_clear(f'\'{sheetName}\'!B2:J1501')

ds = sheet.range(f'B2:J1501')
row = rowStart
# 日時１つを１行分としてデータ生成
for theater in theaters:
    print("*********************************************")
    print(f'*** THEATER: {theater["name"]} ***')

    for movie in theater["movies"]:
        print(f'MOVIE: {movie["name"]}')

        for schedule in movie["schedules"]:
            for d in schedule["date"]:
                # 書き込み対象データ位置
                offset = (row - rowStart) * 9
                if offset >= len(ds):
                    # データ生成終了
                    break

                dt = datetime.datetime.fromtimestamp(
                    int(d), datetime.timezone.utc)
                # 日本のタイムゾーン調整
                dt = dt + datetime.timedelta(hours=9)
                # 終了時刻は目安 ＋１０分
                dtEnd = dt + \
                    datetime.timedelta(minutes=int(movie["time"]) + 10)
                dateStr = dt.strftime('%Y/%m/%d')
                timeStartStr = dt.strftime('%H:%M:%S')
                timeStartStrAbout = dt.strftime('%H時台')
                timeEndStr = dtEnd.strftime('%H:%M:%S')

                # 書き込みデータ
                ds[offset + 0].value = movie["name"]
                ds[offset + 1].value = schedule["type"]
                ds[offset + 2].value = dateStr
                ds[offset + 3].value = timeStartStrAbout
                ds[offset + 4].value = timeStartStr
                ds[offset + 5].value = timeEndStr
                ds[offset + 6].value = movie["time"]
                ds[offset + 7].value = theater["name"]
                ds[offset + 8].value = movie["link"]

                row = row + 1

# 書き込み
sheet.update_cells(ds)
