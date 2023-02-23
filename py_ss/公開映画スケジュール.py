#############################
# カスタム情報
#############################
# 読み書き対象スプレッドシートID
import gspread
import datetime
from bs4 import BeautifulSoup  # BeautifulSoupクラスをインポート
import re
import requests
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


def select(url, selector):
    html = requests.get(url)
    # 取得したHTMLをBeautifulSoupを使ってパースします。
    soup = BeautifulSoup(html.content, "html.parser")
    # パースしたHTMLから特定の要素を抽出します。結果はリストで返ってきます。
    elem = soup.select(selector)

    return elem


def getMoviesFromEiga(url, path):
    # パースしたHTMLから特定の要素を抽出します。結果はリストで返ってきます。
    movies = select(url + path, "main .content-container section[data-title]")
    obj = []
    print("*********************************************")
    print(f'GET INFO from {path}')

    for movie in filter(lambda m: len(m.select("h2 a")) > 0, movies):
        print(f'TRY {movie["data-title"]}')
        m = {
            "name": movie["data-title"],
            "schedules": list(map(lambda t: {
                "type": t.select(".movie-type")[0].get_text("/") if len(t.select(".movie-type")) > 0 else "",
                "date": list(map(lambda d: d["data-time"], t.select("td [data-time]"))),
            }, movie.select(".movie-schedule"))),
            "time": "0",
        }

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


# import json

theaters = list(map(lambda name: {
    "name": name,
    "movies": getMoviesFromEiga(prefURL, paths[name]),
}, paths.keys()))

# print (json.dumps(theaters, ensure_ascii=False))


# 認証のためのコード
# from google.colab import auth
# auth.authenticate_user()
# import gspread
# from google.auth import default
# creds, _ = default()
# gc = gspread.authorize(creds)

gc = gspread.oauth(
    credentials_filename='./client_secrets.json',
)


ss = gc.open_by_key(ssID)
sheet = ss.worksheet(sheetName)

# 既存を削除
ss.values_clear(f'\'{sheetName}\'!B2:J501')

ds = sheet.range(f'B2:J501')

row = rowStart
# 日時１つを１行として書き込み
for theater in theaters:
    print("*********************************************")
    print(f'*** THEATER: {theater["name"]} ***')

    for movie in theater["movies"]:
        print(f'MOVIE: {movie["name"]}')

        for schedule in movie["schedules"]:
            for d in schedule["date"]:
                dt = datetime.datetime.fromtimestamp(
                    int(d), datetime.timezone.utc)
                # 日本のタイムゾーン調整
                dt = dt + datetime.timedelta(hours=9)
                # 終了時刻は目安 ＋１０分
                dtEnd = dt + \
                    datetime.timedelta(minutes=int(movie["time"]) + 10)
                dateStr = dt.strftime('%Y/%m/%d')
                timeStartStr = dt.strftime('%H:%M:%S')
                timeStartStrAbout = dt.strftime('%p%H時台')
                timeEndStr = dtEnd.strftime('%H:%M:%S')

                # 書き込みデータ
                offset = (row - rowStart) * 9
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
