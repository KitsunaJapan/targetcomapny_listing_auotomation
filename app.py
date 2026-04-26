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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
