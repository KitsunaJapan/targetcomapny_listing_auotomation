from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from bs4 import BeautifulSoup
import requests
import time
import os
import re

app = Flask(__name__)
CORS(app)

APP_PASSWORD   = os.environ.get("APP_PASSWORD", "password123")
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-key")

PREF_CODES = {
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

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}


def check_auth():
    return request.headers.get("X-App-Password", "") == APP_PASSWORD


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


@app.route("/api/scrape", methods=["POST"])
def api_scrape():
    """
    サーバー側でNAVITIMEにアクセスしてHTMLをパースする。
    """
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    data      = request.json
    pref_name = data.get("pref_name", "").strip()
    tags      = data.get("tags", "").strip()
    page      = int(data.get("page", 1))

    pref_code = PREF_CODES.get(pref_name)
    if not pref_code:
        return jsonify({"error": f"都道府県名が不正です: {pref_name}"}), 400
    if not tags:
        return jsonify({"error": "業種タグは必須です"}), 400

    url = f"https://www.navitime.co.jp/category/0516/{pref_code}/?page={page}&tags={tags}"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        if resp.status_code == 404:
            return jsonify({"success": True, "fetched": 0, "results": []})
        if not resp.ok:
            return jsonify({"error": f"NAVITIMEアクセスエラー: HTTP {resp.status_code}"}), 400

        today   = time.strftime("%Y-%m-%d")
        results = []
        soup    = BeautifulSoup(resp.text, "html.parser")
        items   = soup.select("li.spot-section")

        for item in items:
            name_el  = item.select_one(".spot-name-text")
            addr_el  = item.select_one(".spot-address")
            link_el  = item.select_one("a[href^='/poi?spot=']")
            tag_els  = item.select(".spot-tag")

            name       = name_el.text.strip()  if name_el else ""
            address    = addr_el.text.strip()   if addr_el else ""
            detail_url = ("https://www.navitime.co.jp" + link_el["href"]) if link_el else ""

            industry_name = ""
            for tag in tag_els:
                t = tag.text.strip().lstrip("#")
                if t:
                    industry_name = t
                    break

            if not name:
                continue

            results.append({
                "取得日":   today,
                "法人名":   name,
                "業種名":   industry_name,
                "電話番号": "",
                "住所":     address,
                "HP":       detail_url,
                "都道府県": pref_name,
            })

        return jsonify({
            "success": True,
            "fetched": len(results),
            "next_page": page + 1,
            "results": results,
        })

    except Exception as e:
        return jsonify({"error": f"スクレイピングエラー: {str(e)}"}), 500


@app.route("/api/write_sheet", methods=["POST"])
def write_sheet():
    """
    OAuthトークンでスプレッドシートに直接書き込む。
    シートがなければ新規作成、あれば追記。重複（法人名）は除外。
    """
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    data       = request.json
    token      = data.get("token", "").strip()
    sheet_id   = data.get("sheet_id", "").strip()
    sheet_name = data.get("sheet_name", "").strip()
    rows       = data.get("rows", [])

    if not token or not sheet_id or not rows:
        return jsonify({"error": "必須パラメータが不足しています"}), 400

    auth_h   = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    meta_url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}"

    # スプレッドシート確認
    meta = requests.get(meta_url, headers={"Authorization": f"Bearer {token}"}, timeout=10)
    if not meta.ok:
        if meta.status_code in (401, 403):
            return jsonify({"error": "Google OAuthトークンが期限切れです。設定から再取得してください。"}), 400
        return jsonify({"error": "スプレッドシートへのアクセスに失敗しました"}), 400

    titles = [s["properties"]["title"] for s in meta.json().get("sheets", [])]
    SHEET_HEADERS = ["取得日", "法人名", "業種名", "電話番号", "メール", "住所", "HP", "都道府県"]

    if sheet_name not in titles:
        # 新規シート作成
        requests.post(f"{meta_url}:batchUpdate", headers=auth_h,
            json={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}, timeout=10)
        # ヘッダー書き込み
        rng = requests.utils.quote(f"{sheet_name}!A1", safe="")
        requests.post(
            f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{rng}:append"
            "?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
            headers=auth_h, json={"values": [SHEET_HEADERS]}, timeout=10)
        mode = "新規作成"
    else:
        mode = "追記"

    # 既存の法人名を取得して重複除外
    rng = requests.utils.quote(f"{sheet_name}!B:B", safe="")
    existing_res = requests.get(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{rng}",
        headers={"Authorization": f"Bearer {token}"}, timeout=10)
    existing_names = set()
    if existing_res.ok:
        for r in existing_res.json().get("values", [])[1:]:
            if r:
                existing_names.add(r[0])

    new_rows = [r for r in rows if r[1] not in existing_names]
    dup_count = len(rows) - len(new_rows)

    if not new_rows:
        return jsonify({"success": True, "written": 0, "duplicates": dup_count, "mode": mode})

    # 一括書き込み
    rng = requests.utils.quote(f"{sheet_name}!A:H", safe="")
    res = requests.post(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{rng}:append"
        "?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
        headers=auth_h, json={"values": new_rows}, timeout=15)

    if not res.ok:
        return jsonify({"error": res.json().get("error", {}).get("message", "書き込みエラー")}), 500

    return jsonify({"success": True, "written": len(new_rows), "duplicates": dup_count, "mode": mode})



# ── Claude API + Web検索でメールアドレスを調査 ────────────
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

@app.route("/api/find_email", methods=["POST"])
def find_email():
    """
    Claude API + Web検索ツールで企業のメールアドレスを調査する。
    1件ずつ処理してメモリを節約する。
    """
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    if not ANTHROPIC_API_KEY:
        return jsonify({"error": "ANTHROPIC_API_KEYが設定されていません"}), 400

    data     = request.json
    name     = data.get("name", "").strip()
    address  = data.get("address", "").strip()
    hp       = data.get("hp", "").strip()

    if not name:
        return jsonify({"error": "法人名は必須です"}), 400

    prompt = f"""以下の企業のメールアドレスを調べてください。

企業名: {name}
住所: {address}
HP: {hp if hp else "不明"}

Web検索で企業の公式HPや問い合わせページを調べ、メールアドレスを見つけてください。
メールアドレスが見つかった場合は、そのメールアドレスのみを返してください。
見つからない場合は「不明」とだけ返してください。
複数ある場合は最も代表的なものを1つだけ返してください。"""

    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "Content-Type": "application/json",
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
            },
            json={
                "model": "claude-haiku-4-5-20251001",
                "max_tokens": 100,
                "tools": [{"type": "web_search_20250305", "name": "web_search"}],
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=30,
        )

        if not resp.ok:
            try:
                err_detail = resp.json()
            except Exception:
                err_detail = resp.text[:300]
            print(f"Claude API error: {resp.status_code} / {err_detail}", flush=True)
            return jsonify({"error": f"Claude APIエラー: {resp.status_code}", "detail": str(err_detail)}), 500

        # レスポンスからテキストを抽出
        content_blocks = resp.json().get("content", [])
        email = ""
        for block in content_blocks:
            if block.get("type") == "text":
                text = block.get("text", "").strip()
                # メールアドレスを正規表現で抽出
                match = re.search(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
                if match:
                    email = match.group()
                break

        return jsonify({"success": True, "email": email})

    except Exception as e:
        return jsonify({"error": f"エラー: {str(e)}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
