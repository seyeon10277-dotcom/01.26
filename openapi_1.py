import streamlit as st
import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
from io import BytesIO
from dotenv import load_dotenv
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 1. 환경 변수 및 OpenAI 설정
load_dotenv()
api_key = os.getenv("Open_api_key")
client = OpenAI(api_key=api_key)

# 2. 페이지 기본 설정
st.set_page_config(page_title="Trade Master 2026", layout="wide", page_icon="🚢")

# Matplotlib 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# 스타일링
st.markdown("""
    <style>
    .main { background-color: #f4f7f9; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    .stDownloadButton>button { width: 100%; background-color: #28a745; color: white; }
    </style>
    """, unsafe_allow_html=True)

# --- [Helper Functions] 데이터 및 서류 생성 ---

@st.cache_data
def get_exchange_data():
    """환율 데이터는 한 번 생성하면 세션 내에서 유지"""
    dates = pd.date_range(end=datetime(2026, 1, 26), periods=30)
    values = [1476, 1478, 1443, 1446, 1441, 1441, 1434, 1437, 1443, 1443,
              1441, 1440, 1445, 1444, 1446, 1450, 1456, 1455, 1464, 1473,
              1462, 1468, 1472, 1471, 1471, 1477, 1464, 1463, 1446, 1439.36]
    return pd.DataFrame({"날짜": dates, "환율": values})

def create_ci_docx(data):
    """상업송장(Commercial Invoice) .docx 생성"""
    doc = Document()
    title = doc.add_heading('COMMERCIAL INVOICE', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table = doc.add_table(rows=5, cols=2)
    table.style = 'Table Grid'
    
    table.rows[0].cells[0].text = f"① Shipper/Seller:\n{data['shipper']}"
    table.rows[0].cells[1].text = f"⑦ Invoice No. and date:\n{data['inv_no_date']}"
    table.rows[1].cells[0].text = f"② Consignee:\n{data['consignee']}"
    table.rows[1].cells[1].text = f"⑧ L/C No. and date:\n{data['lc_no_date']}"
    table.rows[2].cells[0].text = f"⑨ Buyer:\n{data['buyer']}"
    table.rows[2].cells[1].text = f"⑩ Other references:\n{data['other_ref']}"
    table.rows[3].cells[0].text = f"③ Departure date: {data['dep_date']}"
    table.rows[3].cells[1].text = f"⑪ Terms: {data['terms']}"
    table.rows[4].cells[0].text = f"④ Vessel: {data['vessel']} / ⑤ From: {data['from_port']}"
    table.rows[4].cells[1].text = f"⑥ To: {data['to_port']}"

    item_table = doc.add_table(rows=2, cols=6)
    item_table.style = 'Table Grid'
    hdr = item_table.rows[0].cells
    for i, txt in enumerate(['⑫ Marks', '⑬ Pkgs', '⑭ Description', '⑮ Qty', '⑯ Price', '⑰ Amount']):
        hdr[i].text = txt
    
    row = item_table.rows[1].cells
    row[0].text, row[1].text, row[2].text = data['marks'], data['pkg_kind'], data['description']
    row[3].text, row[4].text, row[5].text = str(data['qty']), str(data['unit_price']), str(data['amount'])
    
    doc.add_paragraph(f"\n⑱ Signed by: {data['shipper'].splitlines()[0]}")
    return doc

def create_pl_docx(data):
    """포장명세서(Packing List) .docx 생성"""
    doc = Document()
    doc.add_heading('PACKING LIST', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    table = doc.add_table(rows=4, cols=2)
    table.style = 'Table Grid'
    table.rows[0].cells[0].text = f"① Seller: {data['shipper']}"
    table.rows[0].cells[1].text = f"⑦ Inv No: {data['inv_no_date']}"
    table.rows[1].cells[0].text = f"② Consignee: {data['consignee']}"
    table.rows[1].cells[1].text = f"⑧ Buyer: {data['buyer']}"
    table.rows[2].cells[0].text = f"③ Dep. Date: {data['dep_date']}"
    table.rows[2].cells[1].text = f"⑨ Ref: {data['other_ref']}"
    table.rows[3].cells[0].text = f"④ Vessel: {data['vessel']} / From: {data['from_port']}"
    table.rows[3].cells[1].text = f"⑥ To: {data['to_port']}"

    item_table = doc.add_table(rows=2, cols=6)
    item_table.style = 'Table Grid'
    hdr = item_table.rows[0].cells
    for i, txt in enumerate(['⑩ Marks', '⑪ Pkgs', '⑫ Goods', '⑬ N.W', '⑭ G.W', '⑮ Meas']):
        hdr[i].text = txt
    row = item_table.rows[1].cells
    for i, key in enumerate(['marks', 'pkg_kind', 'description', 'net_weight', 'gross_weight', 'measure']):
        row[i].text = str(data[key])
    return doc

def create_bl_docx(data):
    """선하증권(Bill of Lading) .docx 생성"""
    doc = Document()
    doc.add_heading('BILL OF LADING', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    table = doc.add_table(rows=5, cols=2)
    table.style = 'Table Grid'
    table.rows[0].cells[0].text = f"① Shipper:\n{data['shipper']}"
    table.rows[0].cells[1].text = f"⑪ B/L No: {data['bl_no']}"
    table.rows[1].cells[0].text = f"② Consignee:\n{data['consignee']}"
    table.rows[1].cells[1].text = f"⑦ Voyage: {data['voyage']}"
    table.rows[2].cells[0].text = f"③ Notify:\n{data['notify']}"
    table.rows[2].cells[1].text = f"④ Vessel: {data['vessel']}"
    table.rows[3].cells[0].text = f"⑤ Loading: {data['from_port']}"
    table.rows[3].cells[1].text = f"⑧ Discharge: {data['to_port']}"
    table.rows[4].cells[0].text = f"⑬ Container: {data['cntr_no']}"
    table.rows[4].cells[1].text = f"⑭ Seal No: {data['seal_no']}"
    return doc

# --- [Sidebar] 금융 정보 ---

exchange_df = get_exchange_data()
current_rate = exchange_df["환율"].iloc[-1]
delta = round(current_rate - exchange_df["환율"].iloc[-2], 2)

with st.sidebar:
    st.title("💰 금융 & 실무 현황")
    st.metric(label="USD/KRW (2026-01-26)", value=f"{current_rate:,.2f}원", delta=f"{delta}원")
    st.markdown("---")
    st.subheader("✅ 진행 현황")
    st.checkbox("데이터 입력 완료", key="step1")
    st.checkbox("분석 및 서류 생성 완료", key="step2")

# --- [Main Layout] 상단 대시보드 ---

st.title("🚢 Trade Master 2026: 통합 무역 자동화")

col_graph, col_calc = st.columns([2, 1])
with col_graph:
    st.subheader("📊 환율 추이 (최근 30일)")
    fig, ax = plt.subplots(figsize=(10, 2.5))
    ax.plot(exchange_df["날짜"], exchange_df["환율"], color='#1f77b4', marker='o', markersize=3)
    st.pyplot(fig)

with col_calc:
    st.subheader("🧮 간이 환산기")
    input_usd = st.number_input("달러(USD) 입력", value=1000.0)
    st.write(f"**원화 환산액:** {input_usd * current_rate:,.0f} 원")

# --- [Middle] 정보 입력 폼 ---

st.divider()
st.subheader("📑 거래 및 품목 정보 입력")
with st.form("trade_form"):
    c1, c2 = st.columns(2)
    with c1:
        shipper = st.text_area("수출자(Shipper)", "GILDING TRADING CO., LTD.\nSEOUL, KOREA")
        consignee = st.text_area("수입자(Consignee)", "MONARCH PRO CO., LTD.\nDETROIT, USA")
        vessel = st.text_input("선박/항공편", "PHEONIC")
        from_port = st.text_input("출발지", "BUSAN, KOREA")
        to_port = st.text_input("도착지", "DETROIT, USA")
    with c2:
        inv_no_date = st.text_input("송장번호/날짜", f"INV-2026-001 / {datetime.now().strftime('%b. %d. %Y')}")
        lc_no_date = st.text_input("신용장번호", "LC55352 / APR 25. 2026")
        incoterms = st.selectbox("Incoterms", ["FOB", "EXW", "CIF", "DDP", "CFR"])
        payment = st.selectbox("결제방식", ["L/C AT SIGHT", "T/T PREPAID", "D/P"])
        dep_date = st.text_input("출항일", "MAY. 20. 2026")

    st.markdown("---")
    c3, c4, c5 = st.columns(3)
    description = st.text_input("품명", "NYLON OXFORD")
    with c3:
        qty_input = st.text_input("수량", "60,000M")
        unit_price_input = st.text_input("단가(USD)", "1.00")
        amount_input = st.text_input("총액(USD)", "60,000")
    with c4:
        pkg_kind = st.text_input("포장단위", "53 C/NO")
        net_weight = st.text_input("순중량", "1,200 KGS")
        gross_weight = st.text_input("총중량", "1,208 KGS")
    with c5:
        marks = st.text_input("화인(Marks)", "MON/T DETROIT")
        measure = st.text_input("부피", "5.8 CBM")
        bl_no_input = st.text_input("B/L No", "BK-1004")

    submitted = st.form_submit_button("🚀 분석 및 서류 생성")

# --- [Bottom] 결과 및 다운로드 ---

if submitted:
    data = {
        "shipper": shipper, "consignee": consignee, "vessel": vessel, "from_port": from_port, "to_port": to_port,
        "inv_no_date": inv_no_date, "lc_no_date": lc_no_date, "terms": f"{incoterms} {from_port}", "pay": payment,
        "dep_date": dep_date, "description": description, "qty": qty_input, "unit_price": unit_price_input, 
        "amount": amount_input, "pkg_kind": pkg_kind, "net_weight": net_weight, "gross_weight": gross_weight, 
        "marks": marks, "measure": measure, "bl_no": bl_no_input, "voyage": "1234E", "notify": "SAME AS CONSIGNEE", 
        "cntr_no": "ISCU1104", "seal_no": "S123", "buyer": consignee, "other_ref": "KOREA"
    }
    
    # 세션 상태에 데이터 저장 (재실행 시 유지)
    st.session_state['current_data'] = data

    with st.spinner("AI가 무역 시나리오를 분석 중입니다..."):
        prompt = f"Analyze trade risks for {incoterms} with {qty_input} of {description}. Payment method: {payment}."
        response = client.chat.completions.create(model="gpt-4o", messages=[{"role": "user", "content": prompt}])
        st.session_state['ai_analysis'] = response.choices[0].message.content

if 'ai_analysis' in st.session_state:
    t1, t2 = st.tabs(["💡 AI 전략 가이드", "📥 서류 다운로드"])
    
    with t1:
        st.markdown(st.session_state['ai_analysis'])
        
    with t2:
        st.subheader("📑 자동 생성된 서류 목록")
        current_data = st.session_state['current_data']
        
        doc_files = {
            "Commercial_Invoice.docx": create_ci_docx(current_data),
            "Packing_List.docx": create_pl_docx(current_data),
            "Bill_of_Lading.docx": create_bl_docx(current_data)
        }
        
        cols = st.columns(3)
        for i, (name, doc) in enumerate(doc_files.items()):
            bio = BytesIO()
            doc.save(bio)
            cols[i].download_button(
                label=f"📥 {name}",
                data=bio.getvalue(),
                file_name=name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )