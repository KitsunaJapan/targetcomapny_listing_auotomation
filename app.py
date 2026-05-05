from flask import Flask, request, jsonify, send_from_directory, session
from flask_cors import CORS
import requests
import time
import os

app = Flask(__name__)
CORS(app)

APP_PASSWORD   = os.environ.get("APP_PASSWORD", "password123")
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-key")

GBIZ_API_BASE = "https://info.gbiz.go.jp/hojin/v1"

SHEET_HEADERS = [
    "取得日", "法人番号", "法人名", "業種コード", "業種名",
    "電話番号", "FAX番号", "住所", "HP", "都道府県"
]


def check_auth():
    # X-App-Passwordヘッダーでパスワードを毎回検証（セッション不要）
    pw = request.headers.get("X-App-Password", "")
    return pw == APP_PASSWORD


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/login", methods=["POST"])
def login():
    data = request.json
    if data.get("password") == APP_PASSWORD:
        return jsonify({"success": True})
    return jsonify({"error": "パスワードが違います"}), 401


@app.route("/logout", methods=["POST"])
def logout():
    return jsonify({"success": True})


# ── 企業マスタへの蓄積 ────────────────────────────────────
@app.route("/api/enrich", methods=["POST"])
def api_enrich():
    """
    フロントから法人番号リスト(最大50件/バッチ)を受け取り
    gBizINFO APIで補完して返す。CSVはサーバーに送らない。
    """
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    data       = request.json
    corp_nums  = data.get("corp_nums", [])
    gbiz_token = data.get("gbiz_token", "").strip()
    names      = data.get("names", {})      # {法人番号: 法人名}
    pref_name  = data.get("pref_name", "")  # 都道府県名（CSV由来）

    if not corp_nums or not gbiz_token:
        return jsonify({"error": "法人番号リストとgBizINFO APIトークンは必須です"}), 400

    # メモリ節約のため最大10件に制限（フロント側で10件ずつ送る）
    corp_nums = corp_nums[:10]

    headers = {
        "X-hojinInfo-api-token": gbiz_token,
        "Accept": "application/json",
    }
    today   = time.strftime("%Y-%m-%d")
    results = []

    for corp_num in corp_nums:
        try:
            resp = requests.get(
                f"{GBIZ_API_BASE}/hojin/{corp_num}",
                headers=headers,
                timeout=10,
            )

            if resp.status_code == 404:
                results.append({
                    "取得日":     today,
                    "法人番号":   str(corp_num),
                    "法人名":     names.get(str(corp_num), ""),
                    "業種コード": "",
                    "業種名":     "",
                    "電話番号":   "",
                    "FAX番号":    "",
                    "住所":       "",
                    "HP":         "",
                    "都道府県":   pref_name,
                })
                continue

            if not resp.ok:
                continue

            body = resp.json()
            h    = body.get("hojin-infos", [{}])[0]
            results.append({
                "取得日":     today,
                "法人番号":   str(corp_num),
                "法人名":     h.get("name", "") or names.get(str(corp_num), ""),
                "業種コード": h.get("business_item_number", ""),
                "業種名":     h.get("business_item", ""),
                "電話番号":   h.get("phone_number", ""),
                "FAX番号":    h.get("fax_number", ""),
                "住所":       (h.get("prefecture_name", "") + h.get("city_name", "") + h.get("street_number", "")),
                "HP":         h.get("company_url", ""),
                "都道府県":   h.get("prefecture_name", "") or pref_name,
            })
            # レスポンスを明示的に解放
            del body, h, resp
            time.sleep(0.3)

        except Exception:
            continue

    return jsonify({"success": True, "fetched": len(results), "results": results})


# ── 企業マスタの既存法人番号を取得 ──────────────────────────
@app.route("/api/get_existing_corp_nums", methods=["POST"])
def get_existing_corp_nums():
    """
    企業マスタシートのB列（法人番号）を全件取得して返す。
    フロント側で重複チェックに使う。
    """
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    data     = request.json
    token    = data.get("token", "").strip()
    sheet_id = data.get("sheet_id", "").strip()

    if not token or not sheet_id:
        return jsonify({"error": "token と sheet_id は必須です"}), 400

    auth_h = {"Authorization": f"Bearer {token}"}
    # B列（法人番号）だけ取得（高速・軽量）
    rng = requests.utils.quote("企業マスタ!B:B", safe="")
    res = requests.get(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{rng}",
        headers=auth_h, timeout=15
    )

    if not res.ok:
        # シートがまだ存在しない場合は空リストを返す
        if res.status_code == 400:
            return jsonify({"success": True, "corp_nums": []})
        return jsonify({"error": "企業マスタの読み込みに失敗しました"}), 400

    values = res.json().get("values", [])
    # 1行目はヘッダー（「法人番号」）なのでスキップ
    corp_nums = [row[0] for row in values[1:] if row]
    return jsonify({"success": True, "corp_nums": corp_nums, "count": len(corp_nums)})


# ── スプレッドシート①への書き込み（企業マスタ蓄積） ─────────
@app.route("/api/write_master", methods=["POST"])
def write_master():
    """
    企業マスタシート（スプレッドシート①）への追記。
    既存の法人番号は重複追加しない（フロント側で制御）。
    """
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    data     = request.json
    token    = data.get("token", "").strip()
    sheet_id = data.get("sheet_id", "").strip()
    rows     = data.get("rows", [])

    if not token or not sheet_id or not rows:
        return jsonify({"error": "必須パラメータが不足しています"}), 400

    auth_h   = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    meta_url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}"
    sheet_name = "企業マスタ"

    # シート存在確認
    meta = requests.get(meta_url, headers=auth_h, timeout=10)
    if not meta.ok:
        return jsonify({"error": "スプレッドシートへのアクセスに失敗しました"}), 400

    titles = [s["properties"]["title"] for s in meta.json().get("sheets", [])]
    if sheet_name not in titles:
        # 新規作成＋ヘッダー
        requests.post(f"{meta_url}:batchUpdate", headers=auth_h,
            json={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}, timeout=10)
        rng = requests.utils.quote(f"{sheet_name}!A1", safe="")
        requests.post(
            f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{rng}:append"
            "?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
            headers=auth_h, json={"values": [SHEET_HEADERS]}, timeout=10)

    # 一括追記
    rng = requests.utils.quote(f"{sheet_name}!A:J", safe="")
    res = requests.post(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{rng}:append"
        "?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
        headers=auth_h, json={"values": rows}, timeout=15)

    if not res.ok:
        return jsonify({"error": res.json().get("error", {}).get("message", "書き込みエラー")}), 500

    return jsonify({"success": True, "written": len(rows)})


# ── スプレッドシート①の読み込み ──────────────────────────
@app.route("/api/read_master", methods=["POST"])
def read_master():
    """
    企業マスタから業種コード・都道府県でフィルタして返す。
    """
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    data          = request.json
    token         = data.get("token", "").strip()
    sheet_id      = data.get("sheet_id", "").strip()
    industry_code = data.get("industry_code", "").strip()  # 空なら全業種
    pref_filter   = data.get("pref", "").strip()           # 空なら全国

    if not token or not sheet_id:
        return jsonify({"error": "token と sheet_id は必須です"}), 400

    auth_h = {"Authorization": f"Bearer {token}"}
    rng    = requests.utils.quote("企業マスタ!A:J", safe="")
    res    = requests.get(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{rng}",
        headers=auth_h, timeout=15)

    if not res.ok:
        try:
            err_detail = res.json().get("error", {}).get("message", "")
        except Exception:
            err_detail = res.text[:200]
        if res.status_code in (401, 403):
            return jsonify({"error": f"Google OAuthトークンが無効または期限切れです。設定タブで再取得してください。（詳細: {err_detail}）"}), 400
        return jsonify({"error": f"企業マスタの読み込みに失敗しました。（HTTP {res.status_code}: {err_detail}）"}), 400

    all_rows = res.json().get("values", [])
    if len(all_rows) < 2:
        return jsonify({"error": "企業マスタにデータがありません。先に収集タブで蓄積してください。"}), 400

    # ヘッダー行をスキップしてフィルタ
    # 列順: 取得日(0) 法人番号(1) 法人名(2) 業種コード(3) 業種名(4)
    #       電話(5) FAX(6) 住所(7) HP(8) 都道府県(9)
    filtered = []
    for row in all_rows[1:]:
        while len(row) < 10:
            row.append("")
        ind_match  = not industry_code or row[3] == industry_code
        pref_match = not pref_filter   or row[9] == pref_filter
        if ind_match and pref_match:
            filtered.append({
                "取得日":     row[0],
                "法人番号":   row[1],
                "法人名":     row[2],
                "業種コード": row[3],
                "業種名":     row[4],
                "電話番号":   row[5],
                "FAX番号":    row[6],
                "住所":       row[7],
                "HP":         row[8],
                "都道府県":   row[9],
            })

    return jsonify({"success": True, "total": len(filtered), "rows": filtered})


# ── スプレッドシート②への転記（営業リスト） ───────────────
@app.route("/api/write_sales", methods=["POST"])
def write_sales():
    """
    企業マスタから絞り込んだ結果を営業リストシート（スプレッドシート②）に転記。
    業種名をシート名にする。既存なら追記、なければ新規作成。
    """
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    data       = request.json
    token      = data.get("token", "").strip()
    sheet_id   = data.get("sheet_id", "").strip()
    sheet_name = data.get("sheet_name", "").strip()  # 業種名
    rows       = data.get("rows", [])

    if not token or not sheet_id or not rows:
        return jsonify({"error": "必須パラメータが不足しています"}), 400

    auth_h   = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    meta_url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}"
    sales_headers = ["登録日付", "会社名", "電話番号", "FAX番号", "メールアドレス", "住所", "HP"]

    meta = requests.get(meta_url, headers={**auth_h, "Content-Type": ""}, timeout=10)
    if not meta.ok:
        return jsonify({"error": "スプレッドシートへのアクセスに失敗しました"}), 400

    titles = [s["properties"]["title"] for s in meta.json().get("sheets", [])]
    if sheet_name not in titles:
        requests.post(f"{meta_url}:batchUpdate", headers=auth_h,
            json={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}, timeout=10)
        rng = requests.utils.quote(f"{sheet_name}!A1", safe="")
        requests.post(
            f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{rng}:append"
            "?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
            headers=auth_h, json={"values": [sales_headers]}, timeout=10)
        mode = "新規作成"
    else:
        mode = "追記"

    rng = requests.utils.quote(f"{sheet_name}!A:G", safe="")
    res = requests.post(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{rng}:append"
        "?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
        headers=auth_h, json={"values": rows}, timeout=15)

    if not res.ok:
        return jsonify({"error": res.json().get("error", {}).get("message", "書き込みエラー")}), 500

    return jsonify({"success": True, "written": len(rows), "mode": mode})



# ── NAVITIMEスクレイピング ────────────────────────────────
NAVITIME_PREF_CODES = {
    "北海道":"01","青森県":"02","岩手県":"03","宮城県":"04","秋田県":"05",
    "山形県":"06","福島県":"07","茨城県":"08","栃木県":"09","群馬県":"10",
    "埼玉県":"11","千葉県":"12","東京都":"13","神奈川県":"14","新潟県":"15",
    "富山県":"16","石川県":"17","福井県":"18","山梨県":"19","長野県":"20",
    "岐阜県":"21","静岡県":"22","愛知県":"23","三重県":"24","滋賀県":"25",
    "京都府":"26","大阪府":"27","兵庫県":"28","奈良県":"29","和歌山県":"30",
    "鳥取県":"31","島根県":"32","岡山県":"33","広島県":"34","山口県":"35",
    "徳島県":"36","香川県":"37","愛媛県":"38","高知県":"39","福岡県":"40",
    "佐賀県":"41","長崎県":"42","熊本県":"43","大分県":"44","宮崎県":"45",
    "鹿児島県":"46","沖縄県":"47",
}

@app.route("/api/scrape_navitime", methods=["POST"])
def scrape_navitime():
    """
    NAVITIMEの法人カテゴリページをスクレイピングして企業情報を返す。
    1ページ15件、最大ページ数まで取得。
    """
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    from bs4 import BeautifulSoup

    data      = request.json
    pref_name = data.get("pref_name", "").strip()
    tags      = data.get("tags", "").strip()      # 業種タグ（例: 010429）
    page      = int(data.get("page", 1))          # 取得するページ番号
    max_pages = int(data.get("max_pages", 1))     # 最大ページ数

    pref_code = NAVITIME_PREF_CODES.get(pref_name)
    if not pref_code:
        return jsonify({"error": f"都道府県名が不正です: {pref_name}"}), 400
    if not tags:
        return jsonify({"error": "業種タグは必須です"}), 400

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    }

    today    = time.strftime("%Y-%m-%d")
    results  = []
    end_page = min(page + max_pages - 1, 999)

    for p in range(page, end_page + 1):
        url = f"https://www.navitime.co.jp/category/0516/{pref_code}/?page={p}&tags={tags}"
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            if resp.status_code == 404:
                break
            if not resp.ok:
                return jsonify({"error": f"NAVITIMEアクセスエラー: HTTP {resp.status_code}"}), 400

            soup  = BeautifulSoup(resp.text, "html.parser")
            items = soup.select("li.spot-section")

            if not items:
                break  # これ以上ページがない

            for item in items:
                name_el  = item.select_one(".spot-name-text")
                addr_el  = item.select_one(".spot-address")
                phone_el = item.select_one(".spot-phone-number, .spot-tel, dd.spot-detail-value")
                link_el  = item.select_one("a[href^='/poi?spot=']")
                tag_els  = item.select(".spot-tag")

                name    = name_el.text.strip()  if name_el  else ""
                address = addr_el.text.strip()  if addr_el  else ""
                phone   = phone_el.text.strip() if phone_el else ""
                detail_url = ("https://www.navitime.co.jp" + link_el["href"]) if link_el else ""
                industry = tags  # タグコードを業種コードとして使用

                # タグから業種名を取得
                industry_name = ""
                for tag in tag_els:
                    t = tag.text.strip().lstrip("#")
                    if t and t not in ["オフィスビル", "駅周辺"]:
                        industry_name = t
                        break

                if not name:
                    continue

                results.append({
                    "取得日":     today,
                    "法人番号":   "",
                    "法人名":     name,
                    "業種コード": tags,
                    "業種名":     industry_name,
                    "電話番号":   phone,
                    "FAX番号":    "",
                    "住所":       address,
                    "HP":         detail_url,
                    "都道府県":   pref_name,
                })

            time.sleep(1)  # サーバー負荷軽減

        except Exception as e:
            return jsonify({"error": f"スクレイピングエラー: {str(e)}"}), 500

    # 最終ページ数を確認
    try:
        url  = f"https://www.navitime.co.jp/category/0516/{pref_code}/?page=1&tags={tags}"
        resp = requests.get(url, headers=headers, timeout=15)
        soup = BeautifulSoup(resp.text, "html.parser")
        page_links = soup.select("a.paging-number, .paging a")
        total_pages = max([int(a.text.strip()) for a in page_links if a.text.strip().isdigit()] or [1])
    except Exception:
        total_pages = 1

    return jsonify({
        "success":     True,
        "fetched":     len(results),
        "total_pages": total_pages,
        "results":     results,
    })


@app.route("/api/get_navitime_tags", methods=["GET"])
def get_navitime_tags():
    """
    NAVITIMEのカテゴリツリーAPIから業種タグ一覧を取得する。
    """
    try:
        resp = requests.get(
            "https://www.navitime.co.jp/async/category/tree",
            headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"},
            timeout=10,
        )
        if not resp.ok:
            return jsonify({"error": "タグ一覧の取得に失敗しました"}), 400
        return jsonify(resp.json())
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
