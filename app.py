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


def check_auth():
    return session.get("authenticated") is True


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/login", methods=["POST"])
def login():
    data = request.json
    if data.get("password") == APP_PASSWORD:
        session["authenticated"] = True
        return jsonify({"success": True})
    return jsonify({"error": "パスワードが違います"}), 401


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"success": True})


@app.route("/api/enrich", methods=["POST"])
def api_enrich():
    """
    フロント（ブラウザ）で国税庁CSVを読み込み・業種フィルタ済みの
    法人番号リストを受け取り、gBizINFO APIで連絡先を補完して返す。
    サーバーはCSVを一切扱わないのでメモリ消費なし。
    """
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    data        = request.json
    corp_nums   = data.get("corp_nums", [])   # 法人番号の配列（最大500件）
    gbiz_token  = data.get("gbiz_token", "").strip()
    names       = data.get("names", {})        # {法人番号: 法人名} のマップ（CSV由来）
    addresses   = data.get("addresses", {})    # {法人番号: 住所} のマップ（CSV由来）

    if not corp_nums or not gbiz_token:
        return jsonify({"error": "法人番号リストとgBizINFO APIトークンは必須です"}), 400

    headers = {
        "X-hojinInfo-api-token": gbiz_token,
        "Accept": "application/json",
    }

    results = []
    for corp_num in corp_nums[:500]:
        try:
            resp = requests.get(
                f"{GBIZ_API_BASE}/hojin/{corp_num}",
                headers=headers,
                timeout=10,
            )

            if resp.status_code == 404:
                # gBizINFOに未登録 → CSV由来の情報だけで登録
                results.append({
                    "登録日付":      "",
                    "会社名":        names.get(str(corp_num), ""),
                    "電話番号":      "",
                    "FAX番号":       "",
                    "メールアドレス": "",
                    "住所":          addresses.get(str(corp_num), ""),
                    "HP":            "",
                })
                continue

            if not resp.ok:
                continue  # その他エラーはスキップ

            h = resp.json().get("hojin-infos", [{}])[0]
            results.append({
                "登録日付":      h.get("update_date", ""),
                "会社名":        h.get("name", "") or names.get(str(corp_num), ""),
                "電話番号":      h.get("phone_number", ""),
                "FAX番号":       h.get("fax_number", ""),
                "メールアドレス": h.get("mail", ""),
                "住所":          (h.get("prefecture_name", "") + h.get("city_name", "") + h.get("street_number", ""))
                                  or addresses.get(str(corp_num), ""),
                "HP":            h.get("company_url", ""),
            })

            time.sleep(0.2)  # レート制限対応

        except Exception:
            continue

    return jsonify({
        "success": True,
        "fetched": len(results),
        "results": results,
    })


@app.route("/write_sheet_batch", methods=["POST"])
def write_sheet_batch():
    """OAuthトークンでSheets APIを直接叩く（名刺リーダーと同じ方式）"""
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    data       = request.json
    token      = data.get("token", "").strip()
    sheet_id   = data.get("sheet_id", "").strip()
    sheet_name = data.get("sheet_name", "").strip()
    rows       = data.get("rows", [])

    if not token or not sheet_id or not rows:
        return jsonify({"error": "token / sheet_id / rows は必須です"}), 400

    auth_header = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    meta_url    = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}"

    meta_res = requests.get(meta_url, headers=auth_header, timeout=10)
    if not meta_res.ok:
        err = meta_res.json().get("error", {}).get("message", "スプレッドシートへのアクセスに失敗しました")
        return jsonify({"error": err}), 400

    existing_sheets = [s["properties"]["title"] for s in meta_res.json().get("sheets", [])]

    if sheet_name not in existing_sheets:
        add_res = requests.post(
            f"{meta_url}:batchUpdate",
            headers=auth_header,
            json={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]},
            timeout=10,
        )
        if not add_res.ok:
            return jsonify({"error": "シートの新規作成に失敗しました"}), 500

        header_range = requests.utils.quote(f"{sheet_name}!A1", safe="")
        requests.post(
            f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{header_range}:append"
            "?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
            headers=auth_header,
            json={"values": [["登録日付", "会社名", "電話番号", "FAX番号", "メールアドレス", "住所", "HP"]]},
            timeout=10,
        )
        mode = "新規作成"
    else:
        mode = "追記"

    range_enc = requests.utils.quote(f"{sheet_name}!A:G", safe="")
    write_url = (
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/"
        f"{range_enc}:append"
        "?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS"
    )
    res = requests.post(write_url, headers=auth_header, json={"values": rows}, timeout=15)
    if not res.ok:
        err = res.json().get("error", {}).get("message", "書き込みエラー")
        return jsonify({"error": err}), 500

    return jsonify({"success": True, "written": len(rows), "mode": mode})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
