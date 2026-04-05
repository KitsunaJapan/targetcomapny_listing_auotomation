from flask import Flask, request, jsonify, send_from_directory, session
from flask_cors import CORS
import requests
import time
import os

app = Flask(__name__)
CORS(app)

APP_PASSWORD   = os.environ.get("APP_PASSWORD", "password123")
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-key")

# v2はリニューアル後エンドポイントが変わったため v1 を使用（当面利用可能）
GBIZ_API_BASE = "https://api.info.gbiz.go.jp/hojin/v1"

INDUSTRY_CATEGORIES = {
    "A": "農業、林業",
    "B": "漁業",
    "C": "鉱業、採石業、砂利採取業",
    "D": "建設業",
    "E": "製造業",
    "F": "電気・ガス・熱供給・水道業",
    "G": "情報通信業",
    "H": "運輸業、郵便業",
    "I": "卸売業、小売業",
    "J": "金融業、保険業",
    "K": "不動産業、物品賃貸業",
    "L": "学術研究、専門・技術サービス業",
    "M": "宿泊業、飲食サービス業",
    "N": "生活関連サービス業、娯楽業",
    "O": "教育、学習支援業",
    "P": "医療、福祉",
    "Q": "複合サービス事業",
    "R": "サービス業（他に分類されないもの）",
    "S": "公務（他に分類されるものを除く）",
    "T": "分類不能の産業",
}


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


@app.route("/api/fetch", methods=["POST"])
def api_fetch():
    """gBizINFO APIから業種別企業情報を取得する"""
    if not check_auth():
        return jsonify({"error": "認証が必要です"}), 401

    data          = request.json
    industry_code = data.get("industry_code", "").strip()
    gbiz_token    = data.get("gbiz_token", "").strip()
    max_results   = min(int(data.get("max_results", 50)), 500)

    if not industry_code or not gbiz_token:
        return jsonify({"error": "業種とgBizINFO APIトークンは必須です"}), 400

    # v1 API ヘッダー
    headers   = {
        "X-hojinInfo-api-token": gbiz_token,
        "Accept": "application/json",
    }
    companies = []
    offset    = 1
    limit     = 10

    try:
        while len(companies) < max_results:
            resp = requests.get(
                f"{GBIZ_API_BASE}/hojin",
                headers=headers,
                # v1のパラメータ: category ではなく industry
                params={"industry": industry_code, "limit": limit, "offset": offset},
                timeout=15,
            )
            # エラー詳細をそのまま返してデバッグしやすくする
            if not resp.ok:
                try:
                    err_body = resp.json()
                except Exception:
                    err_body = resp.text
                return jsonify({
                    "error": f"gBizINFO APIエラー (HTTP {resp.status_code})",
                    "detail": err_body,
                }), 400

            body       = resp.json()
            hojin_list = body.get("hojin-infos", [])
            if not hojin_list:
                break

            for h in hojin_list:
                companies.append({
                    "登録日付":      h.get("update_date", ""),
                    "会社名":        h.get("name", ""),
                    "電話番号":      h.get("phone_number", ""),
                    "FAX番号":       h.get("fax_number", ""),
                    "メールアドレス": h.get("mail", ""),
                    "住所":          (h.get("prefecture_name", "") + h.get("city_name", "") + h.get("street_number", "")),
                    "HP":            h.get("company_url", ""),
                })

            total   = body.get("totalCount", 0)
            offset += limit
            time.sleep(0.3)
            if offset > min(total, max_results):
                break

        return jsonify({
            "success":   True,
            "fetched":   len(companies),
            "industry":  INDUSTRY_CATEGORIES.get(industry_code, industry_code),
            "companies": companies,
        })

    except Exception as e:
        return jsonify({"error": f"取得エラー: {str(e)}"}), 500


@app.route("/write_sheet_batch", methods=["POST"])
def write_sheet_batch():
    """
    名刺リーダーと同じ認証方式：
    フロントから渡されたOAuth tokenでSheets APIを直接叩く
    """
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

    # スプレッドシートのメタ情報取得（アクセス確認 + シート一覧）
    meta_res = requests.get(meta_url, headers=auth_header, timeout=10)
    if not meta_res.ok:
        err = meta_res.json().get("error", {}).get("message", "スプレッドシートへのアクセスに失敗しました")
        return jsonify({"error": err}), 400

    existing_sheets = [s["properties"]["title"] for s in meta_res.json().get("sheets", [])]

    # シートがなければ新規作成 → ヘッダー書き込み
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

    # データを一括 append
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
