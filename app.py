# -*- coding: utf-8 -*-
"""
糖尿病・糖尿病予備群向け 総合カウンセリングアプリ（カウンセラー専用）
Streamlit + Gemini API + Google Sheets 連携

【事前準備】
1. pip install streamlit google-generativeai gspread google-auth pandas

2. .streamlit/secrets.toml を以下のように作成してください
--------------------------------------------------
GEMINI_API_KEY = "あなたのGemini APIキー"
SPREADSHEET_ID = "GoogleスプレッドシートのID（URLの/d/と/editの間の文字列）"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "xxxx@xxxx.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
--------------------------------------------------
※ サービスアカウントのメールアドレスを、対象のGoogleスプレッドシートに
   「編集者」権限で共有しておく必要があります。

3. streamlit run app.py で起動
"""

import json
from datetime import datetime, date

import pandas as pd
import streamlit as st

# ---- Gemini ----
import google.generativeai as genai

# ---- Google Sheets ----
import gspread
from google.oauth2.service_account import Credentials

# =====================================================================
# 基本設定
# =====================================================================
st.set_page_config(
    page_title="糖尿病・予備群 総合カウンセリングアプリ（カウンセラー用）",
    layout="wide",
)

SHEET_NAME_RECORDS = "counseling_records"  # 記録用シート名
SHEET_NAME_PATIENTS = "patients"           # 患者マスタ用シート名

STANDARD_LABELS = [
    "HbA1c(%)", "空腹時血糖(mg/dL)", "随時血糖(mg/dL)", "食後2h血糖(mg/dL)",
    "LDLコレステロール(mg/dL)", "HDLコレステロール(mg/dL)", "中性脂肪(mg/dL)",
    "eGFR(mL/min/1.73m2)", "尿蛋白", "AST(U/L)", "ALT(U/L)", "γGTP(U/L)",
    "尿酸(mg/dL)", "収縮期血圧(mmHg)", "拡張期血圧(mmHg)", "体重(kg)", "BMI",
]

ADDITIONAL_LABELS = [
    "亜鉛(μg/dL)", "マグネシウム(mg/dL)", "ビタミンB1(ng/mL)",
    "ビタミンB6(ng/mL)", "ビタミンB12(pg/mL)", "葉酸(ng/mL)",
    "テストステロン(ng/dL)", "游離テストステロン(pg/mL)",
    "TSH(μIU/mL)", "FT3(pg/mL)", "FT4(ng/dL)",
    "フェリチン(ng/mL)", "CPR(ng/mL)（インスリン分泌能）",
    "HOMA-IR", "コルチゾール(μg/dL)",
]

SYMPTOM_LIST = [
    "強い倦怠感・疲労感", "急激な体重減少", "急激な体重増加",
    "多飲・多尿", "手足のしびれ・感覚異常", "視力低下・目のかすみ",
    "眼底出血の指摘歴", "こむら返り・筋けいれん", "筋力低下・サルコペニア様症状",
    "低血糖症状（冷や汗・動悸・手の震え）", "抑うつ・気分の落ち込み",
    "性欲減退・ED症状", "月経異常", "むくみ", "動悸・息切れ", "特に症状なし",
]

MEDICATION_LIST = [
    "GLP-1受容体作動薬（オゼンピック等）", "GIP/GLP-1受容体作動薬（マンジャロ等）",
    "持効型インスリン", "超速効型インスリン", "混合型インスリン",
    "SGLT2阻害薬", "DPP-4阻害薬", "ビグアナイド薬（メトホルミン）",
    "SU薬", "α-グルコシダーゼ阻害薬", "チアゾリジン薬", "その他経口薬",
    "服薬なし",
]

EXERCISE_HISTORY_OPTIONS = [
    "特になし（運動習慣なし）", "学校の体育のみ", "陸上競技（長距離）",
    "陸上競技（短距離・跳躍・投擲）", "水泳", "球技（サッカー・バスケ等）",
    "野球・ソフトボール", "柔道・レスリング等格闘系", "体操・新体操",
    "ラグビー・アメフト", "自転車競技", "その他持久系競技",
    "その他瞬発・筋力系競技",
]

# =====================================================================
# 外部サービス初期化
# =====================================================================
@st.cache_resource(show_spinner=False)
def init_gemini_model():
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        return None
    genai.configure(api_key=api_key)
    return genai.GenerativeModel("gemini-3.6-flash")



@st.cache_resource(show_spinner=False)
def init_gsheet_client():
    if "gcp_service_account" not in st.secrets:
        return None
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scopes
    )
    return gspread.authorize(creds)


def get_or_create_worksheet(sh, title, header):
    try:
        ws = sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=2000, cols=max(20, len(header)))
        ws.append_row(header)
    return ws


def get_spreadsheet():
    client = init_gsheet_client()
    sheet_id = st.secrets.get("SPREADSHEET_ID", "")
    if client is None or not sheet_id:
        return None
    return client.open_by_key(sheet_id)


# =====================================================================
# データ保存・読み込み
# =====================================================================
RECORD_HEADER = [
    "timestamp", "patient_id", "patient_name", "birthdate", "gender",
    "exercise_history", "standard_data_json", "additional_data_json",
    "symptoms_json", "medications_json", "counselor_analysis", "patient_summary",
]

PATIENT_HEADER = ["patient_id", "patient_name", "birthdate", "gender", "exercise_history", "memo"]


def save_record_to_sheet(record: dict):
    sh = get_spreadsheet()
    if sh is None:
        st.warning("Googleスプレッドシートが未接続のため、記録は保存されませんでした（secrets.tomlを確認してください）。")
        return False
    ws = get_or_create_worksheet(sh, SHEET_NAME_RECORDS, RECORD_HEADER)
    row = [record.get(k, "") for k in RECORD_HEADER]
    ws.append_row(row)
    return True


def save_patient_to_sheet(patient: dict):
    sh = get_spreadsheet()
    if sh is None:
        return False
    ws = get_or_create_worksheet(sh, SHEET_NAME_PATIENTS, PATIENT_HEADER)
    values = ws.get_all_records()
    existing_ids = [str(v.get("patient_id")) for v in values]
    if str(patient["patient_id"]) in existing_ids:
        idx = existing_ids.index(str(patient["patient_id"])) + 2  # header分+1-index
        ws.update(f"A{idx}:F{idx}", [[patient.get(k, "") for k in PATIENT_HEADER]])
    else:
        ws.append_row([patient.get(k, "") for k in PATIENT_HEADER])
    return True


@st.cache_data(ttl=30, show_spinner=False)
def load_patients_df():
    sh = get_spreadsheet()
    if sh is None:
        return pd.DataFrame(columns=PATIENT_HEADER)
    ws = get_or_create_worksheet(sh, SHEET_NAME_PATIENTS, PATIENT_HEADER)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=PATIENT_HEADER)
    return pd.DataFrame(data)


@st.cache_data(ttl=30, show_spinner=False)
def load_records_df():
    sh = get_spreadsheet()
    if sh is None:
        return pd.DataFrame(columns=RECORD_HEADER)
    ws = get_or_create_worksheet(sh, SHEET_NAME_RECORDS, RECORD_HEADER)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=RECORD_HEADER)
    return pd.DataFrame(data)


def get_patient_history(patient_id: str) -> pd.DataFrame:
    df = load_records_df()
    if df.empty or "patient_id" not in df.columns:
        return df
    return df[df["patient_id"].astype(str) == str(patient_id)].sort_values("timestamp")


# =====================================================================
# Gemini プロンプト構築
# =====================================================================
def build_prompt(patient, standard_data, additional_data, symptoms, medications, history_df):
    """
    カウンセラー向けAI解析用プロンプトを構築する。
    ロジックのポイント：
    - HbA1c/血糖値の「数値が下がった＝良い」という単純評価を避け、
      降下速度・急激さを重視し、眼底出血リスクや筋肉量低下（サルコペニア化）を
      あわせて評価するよう明示的に指示する。
    - 成長期（〜20歳まで）の運動歴を踏まえ、スポーツ心臓（心肥大）や
      高い基礎筋肉量を持つ被検者では、一般的な基準値の読み替えが必要になる
      ことを考慮させる。
    - 亜鉛・マグネシウム・ビタミンB群・テストステロン・甲状腺機能などの
      追加項目について、現在のデータで不足していれば追加オーダーを提案し、
      ミトコンドリア活性化に資する補酵素・微量栄養素の提案も行わせる。
    """

    history_text = "特になし（初回、または過去データなし）"
    if history_df is not None and not history_df.empty:
        rows = []
        for _, r in history_df.iterrows():
            try:
                sd = json.loads(r.get("standard_data_json", "{}") or "{}")
            except Exception:
                sd = {}
            rows.append(f"{r.get('timestamp')}: {json.dumps(sd, ensure_ascii=False)}")
        history_text = "\n".join(rows[-10:])  # 直近10件

    prompt = f"""
あなたは糖尿病・糖尿病予備群の患者を担当する、経験豊富な臨床栄養・内分泌領域の
専門カウンセラーを支援するAIアシスタントです。以下の情報をもとに、
担当カウンセラーが患者指導に使うための分析を行ってください。

# 患者基本情報
氏名: {patient.get('patient_name')}
生年月日: {patient.get('birthdate')}
性別: {patient.get('gender')}
成長期（〜20歳まで）の運動歴: {patient.get('exercise_history')}

# 過去の検査値推移（時系列、直近10件）
{history_text}

# 今回の検査データ（通常項目）
{json.dumps(standard_data, ensure_ascii=False, indent=2)}

# 今回の検査データ（追加オーダー項目）
{json.dumps(additional_data, ensure_ascii=False, indent=2)}

# 自覚症状チェックリスト（今回申告分）
{json.dumps(symptoms, ensure_ascii=False)}

# 使用中の薬剤
{json.dumps(medications, ensure_ascii=False)}

# 分析にあたって必ず考慮すべき観点

1. 【降下速度と急変リスクの評価】
   HbA1cや血糖値は「下がっていれば良い」と単純評価せず、前回値との差分から
   降下スピードを計算し、月あたりの低下幅が急激（目安として月1.5〜2%を超える
   HbA1c低下や、短期間での血糖の大幅是正）である場合には、
   - 糖尿病網膜症の急速進行・眼底出血リスクの上昇
   - 除脂肪体重・筋肉量の急速な低下（サルコペニア化、フレイル化）
   の可能性を明示的に指摘し、降下ペースを緩めるための指導方針（薬剤調整の
   相談、たんぱく質摂取、レジスタンス運動の併用など）を提案してください。
   数値の改善自体を安易に「順調」と評価しないこと。

2. 【成長期の運動歴を踏まえたデータの読み替え】
   成長期（〜20歳まで）に持久系競技（長距離走・水泳等）や瞬発・筋力系競技
   （投擲・格闘技・ラグビー等）に本格的に取り組んでいた場合、
   - スポーツ心臓（生理的心肥大）に由来する心電図・心エコー所見の解釈の違い
   - 一般成人平均より高い基礎筋肉量・除脂肪体重を持つことによる
     クレアチニン、eGFR、尿酸値などの解釈の違い
   - 筋肉量の多さが体重・BMIの評価を歪める可能性
   を考慮し、一般的な基準値をそのまま当てはめるのではなく、
   その被検者の運動歴を踏まえた読み替えを行った上でコメントしてください。

3. 【追加項目の必要性評価とミトコンドリア活性化の提案】
   現在のデータ（特に追加オーダー項目）を確認し、
   - 亜鉛、マグネシウム、ビタミンB群（B1, B6, B12, 葉酸）
   - テストステロン（遊離型含む）
   - 甲状腺機能（TSH, FT3, FT4）
   についてデータが不足している場合は、次回オーダーすべき項目として
   具体的に提案してください。
   また、既存データから微量栄養素欠乏や代謝機能低下が疑われる場合は、
   ミトコンドリア機能活性化の観点から、補酵素Q10、αリポ酸、L-カルニチン、
   ビタミンB群、マグネシウム等、エビデンスに基づいた栄養サプリメント・
   補酵素の提案を行ってください（薬剤ではなく栄養指導の範囲であることを
   明記し、処方判断は医師が行う旨を必ず添えること）。

# 出力形式
以下の2パートに分けて、必ず指定の見出しで出力してください。

## カウンセラー用詳細分析
（専門用語を用いてよい。数値の解釈根拠、リスク評価、次回オーダー推奨項目、
栄養・運動・薬剤調整に関する具体的な指導方針を、箇条書き中心にまとめてください）

## 患者提示用ビジュアル解説
（平易な言葉で、患者本人にそのまま見せられるトーンで。専門用語は避けるか
補足を添え、不安を過度に煽らず、前向きな行動指針を中心にまとめてください。
250〜400字程度を目安に簡潔にまとめること）
"""
    return prompt


def call_gemini(prompt: str) -> str:
    model = init_gemini_model()
    if model is None:
        return "【エラー】Gemini APIキーが設定されていません。st.secretsのGEMINI_API_KEYを確認してください。"
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"【エラー】AI解析中に例外が発生しました: {e}"


def split_analysis(full_text: str):
    """AI出力をカウンセラー用/患者用の2パートに分割する"""
    counselor_part, patient_part = full_text, ""
    marker = "## 患者提示用ビジュアル解説"
    if marker in full_text:
        parts = full_text.split(marker)
        counselor_part = parts[0].replace("## カウンセラー用詳細分析", "").strip()
        patient_part = parts[1].strip()
    return counselor_part, patient_part


# =====================================================================
# セッション状態初期化
# =====================================================================
if "selected_patient" not in st.session_state:
    st.session_state.selected_patient = {
        "patient_id": "",
        "patient_name": "",
        "birthdate": date(2000, 1, 1),
        "gender": "未選択",
        "exercise_history": [],
        "memo": "",
    }
if "ai_counselor_text" not in st.session_state:
    st.session_state.ai_counselor_text = ""
if "ai_patient_text" not in st.session_state:
    st.session_state.ai_patient_text = ""

st.title("🩺 糖尿病・糖尿病予備群 総合カウンセリングアプリ（カウンセラー専用）")
st.caption("本アプリはカウンセラー・医療専門職の指導補助を目的としたツールです。最終的な診断・処方判断は必ず医師が行ってください。")

col_left, col_center, col_right = st.columns([1.1, 1.4, 1.5])

# =====================================================================
# 左カラム：患者検索・基本プロフィール
# =====================================================================
with col_left:
    st.subheader("👤 患者検索・基本プロフィール")

    patients_df = load_patients_df()
    search_name = st.text_input("患者名・IDで検索", "")

    if not patients_df.empty:
        if search_name:
            mask = (
                patients_df["patient_name"].astype(str).str.contains(search_name, case=False, na=False)
                | patients_df["patient_id"].astype(str).str.contains(search_name, case=False, na=False)
            )
            filtered = patients_df[mask]
        else:
            filtered = patients_df

        options = ["新規患者を登録"] + [
            f"{row['patient_id']} - {row['patient_name']}" for _, row in filtered.iterrows()
        ]
        choice = st.selectbox("患者を選択", options)

        if choice != "新規患者を登録":
            pid = choice.split(" - ")[0]
            row = patients_df[patients_df["patient_id"].astype(str) == pid].iloc[0]
            st.session_state.selected_patient = {
                "patient_id": row.get("patient_id", ""),
                "patient_name": row.get("patient_name", ""),
                "birthdate": row.get("birthdate", ""),
                "gender": row.get("gender", "未選択"),
                "exercise_history": str(row.get("exercise_history", "")).split("、") if row.get("exercise_history") else [],
                "memo": row.get("memo", ""),
            }
    else:
        st.info("登録済み患者がまだいません。下記フォームから新規登録してください。")

    st.markdown("---")
    st.markdown("**基本プロフィール入力／編集**")

    sp = st.session_state.selected_patient
    patient_id = st.text_input("患者ID（未入力の場合は自動採番）", value=str(sp.get("patient_id", "")))
    patient_name = st.text_input("氏名", value=sp.get("patient_name", ""))
    gender = st.selectbox(
        "性別", ["未選択", "男性", "女性", "その他・回答しない"],
        index=["未選択", "男性", "女性", "その他・回答しない"].index(sp.get("gender", "未選択"))
        if sp.get("gender", "未選択") in ["未選択", "男性", "女性", "その他・回答しない"] else 0,
    )
    birthdate_default = date(2000, 1, 1)
    try:
        if sp.get("birthdate") and isinstance(sp.get("birthdate"), str):
            birthdate_default = datetime.strptime(sp["birthdate"], "%Y-%m-%d").date()
    except Exception:
        pass
    birthdate = st.date_input("生年月日", value=birthdate_default, min_value=date(1920, 1, 1), max_value=date.today())

    st.markdown("**成長期（〜20歳まで）の運動歴**")
    exercise_history = st.multiselect(
        "該当するものをすべて選択（心肥大・高筋肉量の評価に使用）",
        EXERCISE_HISTORY_OPTIONS,
        default=[e for e in sp.get("exercise_history", []) if e in EXERCISE_HISTORY_OPTIONS],
    )
    memo = st.text_area("メモ（既往歴・特記事項など）", value=sp.get("memo", ""), height=80)

    if st.button("💾 患者プロフィールを保存", use_container_width=True):
        if not patient_name:
            st.error("氏名を入力してください。")
        else:
            final_id = patient_id if patient_id else f"P{int(datetime.now().timestamp())}"
            patient_record = {
                "patient_id": final_id,
                "patient_name": patient_name,
                "birthdate": birthdate.strftime("%Y-%m-%d"),
                "gender": gender,
                "exercise_history": "、".join(exercise_history),
                "memo": memo,
            }
            ok = save_patient_to_sheet(patient_record)
            st.session_state.selected_patient = patient_record
            load_patients_df.clear()
            if ok:
                st.success(f"患者プロフィールを保存しました（ID: {final_id}）")
            else:
                st.warning("スプレッドシート未接続のため、今回のセッションのみで保持されます。")

    st.markdown("---")
    st.markdown("**過去の数値比較（HbA1cの推移）**")
    current_pid = st.session_state.selected_patient.get("patient_id", "")
    if current_pid:
        hist_df = get_patient_history(current_pid)
        if not hist_df.empty:
            chart_rows = []
            for _, r in hist_df.iterrows():
                try:
                    sd = json.loads(r.get("standard_data_json", "{}") or "{}")
                    hba1c = sd.get("HbA1c(%)")
                    if hba1c not in (None, ""):
                        chart_rows.append({"日時": r.get("timestamp"), "HbA1c(%)": float(hba1c)})
                except Exception:
                    continue
            if chart_rows:
                chart_df = pd.DataFrame(chart_rows).set_index("日時")
                st.line_chart(chart_df)
            else:
                st.caption("過去のHbA1cデータがまだありません。")
        else:
            st.caption("この患者の過去記録はまだありません。")
    else:
        st.caption("患者プロフィールを保存すると、過去データ比較が表示されます。")

# =====================================================================
# 中央カラム：検査データ入力・症状・薬剤
# =====================================================================
with col_center:
    st.subheader("🧪 血液検査データ・症状・使用薬剤")

    st.markdown("**通常項目**")
    standard_data = {}
    std_cols = st.columns(2)
    for i, label in enumerate(STANDARD_LABELS):
        with std_cols[i % 2]:
            standard_data[label] = st.text_input(label, key=f"std_{label}")

    with st.expander("➕ 追加オーダー項目（亜鉛・Mg・ビタミンB群・テストステロン・甲状腺機能 等）"):
        additional_data = {}
        add_cols = st.columns(2)
        for i, label in enumerate(ADDITIONAL_LABELS):
            with add_cols[i % 2]:
                additional_data[label] = st.text_input(label, key=f"add_{label}")

    st.markdown("---")
    st.markdown("**自覚症状チェックリスト**")
    symptoms = st.multiselect("該当する症状をすべて選択", SYMPTOM_LIST, key="symptoms_select")

    st.markdown("**使用中の薬剤**")
    medications = st.multiselect("該当する薬剤をすべて選択", MEDICATION_LIST, key="medications_select")
    med_note = st.text_input("薬剤に関する補足（用量・服用開始時期など）", key="med_note")

# =====================================================================
# 右カラム：AI解析結果
# =====================================================================
with col_right:
    st.subheader("🤖 AI解析結果")

    run_col, save_col = st.columns(2)
    with run_col:
        run_analysis = st.button("🔍 AI解析を実行", use_container_width=True, type="primary")
    with save_col:
        save_record = st.button("📤 記録をスプレッドシートへ保存", use_container_width=True)

    if run_analysis:
        current_patient = st.session_state.selected_patient
        if not current_patient.get("patient_name"):
            st.error("先に左カラムで患者プロフィールを入力・保存してください。")
        else:
            with st.spinner("Gemini APIで解析中..."):
                hist_df = get_patient_history(current_patient.get("patient_id", ""))
                prompt = build_prompt(
                    current_patient, standard_data, additional_data,
                    symptoms, medications + ([med_note] if med_note else []),
                    hist_df,
                )
                result_text = call_gemini(prompt)
                counselor_text, patient_text = split_analysis(result_text)
                st.session_state.ai_counselor_text = counselor_text
                st.session_state.ai_patient_text = patient_text or result_text

    tab1, tab2 = st.tabs(["📋 カウンセラー用詳細分析", "🗣️ 患者提示用ビジュアル解説"])
    with tab1:
        if st.session_state.ai_counselor_text:
            st.markdown(st.session_state.ai_counselor_text)
        else:
            st.info("「AI解析を実行」を押すと、ここに詳細分析が表示されます。")

    with tab2:
        if st.session_state.ai_patient_text:
            st.markdown(st.session_state.ai_patient_text)
            st.caption("※このタブの内容はそのまま患者向けに提示できるよう平易な表現にしています。")
        else:
            st.info("「AI解析を実行」を押すと、ここに患者向け解説が表示されます。")

    if save_record:
        current_patient = st.session_state.selected_patient
        if not current_patient.get("patient_name"):
            st.error("先に左カラムで患者プロフィールを入力・保存してください。")
        elif not st.session_state.ai_counselor_text and not st.session_state.ai_patient_text:
            st.warning("AI解析結果がまだありません。先に「AI解析を実行」してください（解析なしでも保存は可能ですが、結果欄は空欄になります）。")
            record = {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "patient_id": current_patient.get("patient_id", ""),
                "patient_name": current_patient.get("patient_name", ""),
                "birthdate": str(current_patient.get("birthdate", "")),
                "gender": current_patient.get("gender", ""),
                "exercise_history": "、".join(current_patient.get("exercise_history", [])) if isinstance(current_patient.get("exercise_history"), list) else current_patient.get("exercise_history", ""),
                "standard_data_json": json.dumps(standard_data, ensure_ascii=False),
                "additional_data_json": json.dumps(additional_data, ensure_ascii=False),
                "symptoms_json": json.dumps(symptoms, ensure_ascii=False),
                "medications_json": json.dumps(medications, ensure_ascii=False),
                "counselor_analysis": st.session_state.ai_counselor_text,
                "patient_summary": st.session_state.ai_patient_text,
            }
            if save_record_to_sheet(record):
                load_records_df.clear()
                st.success("記録を保存しました。")
        else:
            record = {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "patient_id": current_patient.get("patient_id", ""),
                "patient_name": current_patient.get("patient_name", ""),
                "birthdate": str(current_patient.get("birthdate", "")),
                "gender": current_patient.get("gender", ""),
                "exercise_history": "、".join(current_patient.get("exercise_history", [])) if isinstance(current_patient.get("exercise_history"), list) else current_patient.get("exercise_history", ""),
                "standard_data_json": json.dumps(standard_data, ensure_ascii=False),
                "additional_data_json": json.dumps(additional_data, ensure_ascii=False),
                "symptoms_json": json.dumps(symptoms, ensure_ascii=False),
                "medications_json": json.dumps(medications, ensure_ascii=False),
                "counselor_analysis": st.session_state.ai_counselor_text,
                "patient_summary": st.session_state.ai_patient_text,
            }
            if save_record_to_sheet(record):
                load_records_df.clear()
                st.success("記録をスプレッドシートに保存しました。")

st.markdown("---")
st.caption(
    "免責事項: 本アプリのAI解析結果は診断ではなく、カウンセラーによる指導のための参考情報です。"
    "薬剤調整・治療方針の最終決定は必ず医師の判断のもとで行ってください。"
)
