import urllib.parse
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
SS_ID = '11HxD-a2Sq8DXPEEiVfZmNtEqkdOxJKZckPRkkT0GOJY'
# 読み書き対象スプレッドシート シート名
SHEET_NAME_ROADSHOW = '映画館'
SHEET_NAME_COMMING = '公開予定'
# データ表が始まる行
ROW_START = 2

# 検索対象サイト
baseURL = 'https://eiga.com'
# 対象の県
prefURL = f'{baseURL}/theater/10'
paths = {
    "ユナイテッド・シネマ 前橋": "/100101/3229/",
    "109シネマズ高崎": "/100201/3230/",
    "イオンシネマ高崎": "/100201/3231/",
}

# 公開予定
commingURL = 'https://www.unitedcinemas.jp/maebashi/movie.php'


#############################
# 関数
#############################

# 映画名で検索

def searchMovie(movieName):
    # 検索用の補正
    movieName = movieName.replace(".", " ")
    movieName = movieName.replace("/", " ")
    result = select(
        f'{baseURL}/search/{urllib.parse.quote(movieName)}/', "#rslt-movie ul a")
    # ID（を含むリンク）を取得
    movieIds = list(map(lambda html: [html.select("p[class='title']")[
                    0].get_text(), html.get('href')], result))
    if len(movieIds) == 0:
        print(f'★★★注意★★★　[{movieName}]という映画が見つかりませんでした ★★★')
        return {
            'name': movieName,
            'link': '',
        }
    if len(movieIds) > 1:
        print(
            f'★★★注意★★★　[{movieName}]という映画で複数ヒットしました。最初にヒットした情報を設定します（候補：{", ".join(map(lambda l: l[0], movieIds))}）★★★')

    movieHtml = select(f'{baseURL}{movieIds[0][1]}', "main")[0]
    images = movieHtml.select(".icon-movie-poster img")
    infos = movieHtml.select(".movie-details .data")
    directors = movieHtml.select(".movie-staff [itemprop='director']")
    casts = movieHtml.select(".movie-cast .person [itemprop='name']")

    return {
        'name': movieIds[0][0],
        'link': f'{baseURL}{movieIds[0][1]}',
    }

# URLから対象selectorを取得


def select(url, selector):
    html = requests.get(url)
    # 取得したHTMLをBeautifulSoupを使ってパースします。
    soup = BeautifulSoup(html.content, "html.parser")
    # パースしたHTMLから特定の要素を抽出します。結果はリストで返ってきます。
    elem = soup.select(selector)

    return elem

# 1つの映画館ごとの映画情報取得


def getMoviesFromEiga(url, path):
    obj = []
    print("*********************************************")
    print(f'GET INFO from {path}')

    # -2024
    for movie in filter(lambda m: len(m.select("h2 a")) > 0, select(url + path, "main .content-container section[data-title]")):
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

    # 2025-
    for movie in select(url + path, ".main-column #tab01_content.tabs__content > .mb-20"):
        # 1つの映画ごとの情報を取得
        moviename = movie.select(".line-clamp-1")[0].get_text()
        print(f'TRY {moviename}')
        #  .mb-40 => 1つの映画のスケジュール表
        m = {
            "name": moviename,
            "schedules": list(map(lambda t: {
                "type": '/'.join(list(map(lambda ty: ty.get_text(), t.select(".movie-type")))),
                "date": list(),
            }, movie.select(".mb-40"))),
            "time": "0",
        }

        dateStr = str(datetime.date.today())
        print("  対象日付 > " + dateStr)

        # 文字列から時間を生成
        for span in movie.select(".weekly-schedule__btn"):
            # span 直下の文字列が日付の場合のみ追加
            timeStr = str(span.get_text()).strip()
            print("***TEST***" + timeStr)
            if re.match("^(\d+):(\d+)", timeStr):
                timeStr = re.sub("～.*", "", timeStr).strip()
                print("    対象時刻(追加分) > " + timeStr)
                dt = datetime.datetime.strptime(
                    dateStr + " " + timeStr + ' +0900', '%Y-%m-%d %H:%M %z')
                # TODO: schedulesはlistである必要なさそう？
                m["schedules"][0]["date"].append(int(dt.timestamp()))

        print(
            f'  => GET {"/".join(list(map(str, m["schedules"][0]["date"])))}')

        # if len(movie.select(".movie-image img")) > 0:
        #     m["image"] = movie.select(".movie-image img")[0]["src"]

        for i in movie.select(".link-btn-discription"):
            timeResult = re.match("(\d+)[分]*", i.get_text().strip())
            if timeResult:
                m["time"] = timeResult.group(1)

        if len(movie.select("a.w100per")) > 0:
            m["link"] = baseURL + movie.select("a.w100per")[0]["href"]

        obj.append(m)

    return obj


# 公開中映画情報
def getRoadShow(ssID, sheetName, rowStart):
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


def getCommingSoon(ssID, sheetName, rowStart):
    ##### 上映予定映画情報一覧を取得 #####
    movies = getCommingFromEiga(commingURL)

    ##### 取得した情報をSpreadSheetsに出力 #####
    ss = gc.open_by_key(ssID)
    sheet = ss.worksheet(sheetName)

    # 既存を削除
    ss.values_clear(f'\'{sheetName}\'!B2:E1501')

    ds = sheet.range(f'B2:E1501')
    row = rowStart
    for movie in movies:
        # 書き込み対象データ位置
        offset = (row - rowStart) * 4
        if offset >= len(ds):
            # データ生成終了
            break

        # 書き込みデータ
        ds[offset + 0].value = movie["name"]
        ds[offset + 1].value = movie["date"]
        ds[offset + 2].value = movie["info"]
        ds[offset + 3].value = movie["link"]

        row = row + 1

    # 書き込み
    sheet.update_cells(ds)


def getCommingFromEiga(url):
    movies = select(url, "#top #showingList li .movieHead")
    obj = []
    date = ''
    print("*********************************************")
    print(f'GET COMMING INFO from {url}')

    for movie in movies:
        if len(movie.select("h3 strong a")) > 0:
            name = movie.select("h3 strong a")[0].get_text()
            date = movie.select("h3 em")[0].get_text()
            # 公開 以降を削除
            date = re.sub("(（|公開).+", "", date)
            # eiga.comで検索して正規化
            eigaObj = searchMovie(name)
            m = {
                "name": eigaObj["name"],
                "date": date,
                "info": ", ".join(list(map(lambda t: t.get_text(), movie.select(".main > a > p")))),
                "link": eigaObj["link"],
            }
            obj.append(m)
            print(f'映画情報 {name}')
            print(f'公開日付 {date}')

    return obj


#############################
# メイン処理開始
#############################
getRoadShow(SS_ID, SHEET_NAME_ROADSHOW, ROW_START)

getCommingSoon(SS_ID, SHEET_NAME_COMMING, ROW_START)
