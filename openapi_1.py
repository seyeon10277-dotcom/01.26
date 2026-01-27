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
    .info-box { background-color: #e9f5ff; padding: 15px; border-radius: 10px; border-left: 5px solid #007bff; }
    </style>
    """, unsafe_allow_html=True)

# --- [Helper Functions] 데이터 및 서류 생성 ---

@st.cache_data
def get_exchange_data():
    """환율 데이터 생성 (2026-01-26 기준 최근 30일)"""
    dates = pd.date_range(end=datetime(2026, 1, 26), periods=30)
    values = [1476, 1478, 1443, 1446, 1441, 1441, 1434, 1437, 1443, 1443,
              1441, 1440, 1445, 1444, 1446, 1450, 1456, 1455, 1464, 1473,
              1462, 1468, 1472, 1471, 1471, 1477, 1464, 1463, 1446, 1439.36]
    return pd.DataFrame({"날짜": dates, "환율": values})

def calculate_estimated_cost(base_price, term, transport, insurance, payment, fta_type):
    """
    인코텀즈, 운송수단, 보험, 결제방식 및 FTA 관세 혜택을 포함한 최종 금액 산출
    """
    total = base_price
    
    # 1. 운송비 가산 (C, D 조건일 때)
    if term in ["CFR", "CIF", "CPT", "CIP", "DAP", "DPU", "DDP"]:
        freight_rate = 0.15 if transport == "항공(AIR)" else 0.05
        total += base_price * freight_rate
        
    # 2. 보험료 가산
    insurance_rates = {
        "ICC(A) (=ICC(AIR))": 0.008, 
        "ICC(B) (=ICC(WA))": 0.005, 
        "ICC(C) (=ICC(FPA))": 0.003, 
        "ICC(WAIOP)": 0.004,
        "선택 안함": 0
    }
    total += base_price * insurance_rates.get(insurance, 0)
    
    # 3. 결제 방식별 금융 비용
    payment_fees = {
        "사전 송금 (Advance Payment)": 0.0,
        "사후 송금 (Settlement after Shipment)": 0.0,
        "D/P (Documents against Payment)": 0.0015,
        "D/A (Documents against Acceptance)": 0.0025,
        "일람출급 신용장 (Sight L/C)": 0.008,
        "기한부 신용장 (Usance L/C)": 0.015
    }
    total += base_price * payment_fees.get(payment, 0)
    
    # 4. FTA 관세 적용 (DDP 조건일 때 세금 시뮬레이션)
    # 기본 관세 8% + 부가세 10% = 약 18%가 협정 미적용 기준
    fta_rates = {
        "협정 미적용 (기본세율)": 0.18,
        "한-미 FTA (KOR-USA)": 0.10,   # 관세 0% 가정 + 부가세 10%
        "한-EU FTA (KOR-EU)": 0.10,    # 관세 0% 가정 + 부가세 10%
        "한-중 FTA (KOR-CHINA)": 0.14, # 민감품목 관세 일부 잔존 가정
        "한-베트남 FTA (KOR-VIETNAM)": 0.10,
        "RCEP (역내포괄적경제동반자협정)": 0.12
    }
    
    if term == "DDP":
        total += base_price * fta_rates.get(fta_type, 0.18)
        
    return total

def create_ci_docx(data):
    """상업송장(Commercial Invoice) .docx 생성"""
    doc = Document()
    title = doc.add_heading('COMMERCIAL INVOICE', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table = doc.add_table(rows=6, cols=2); table.style = 'Table Grid'
    
    table.rows[0].cells[0].text = f"① Shipper/Seller:\n{data['shipper']}"
    table.rows[0].cells[1].text = f"⑦ Invoice No. and date:\n{data['inv_no_date']}"
    table.rows[1].cells[0].text = f"② Consignee:\n{data['consignee']}"
    table.rows[1].cells[1].text = f"⑧ L/C No. and date:\n{data['lc_no_date']}"
    table.rows[2].cells[0].text = f"⑨ Buyer:\n{data['buyer']}"
    table.rows[2].cells[1].text = f"⑪ Terms: {data['terms']} / {data['transport']}"
    table.rows[3].cells[0].text = f"③ Departure date: {data['dep_date']}"
    table.rows[3].cells[1].text = f"⑫ Insurance: {data['insurance']}"
    table.rows[4].cells[0].text = f"④ Vessel: {data['vessel']} / From: {data['from_port']}"
    table.rows[4].cells[1].text = f"⑥ To: {data['to_port']}"
    table.rows[5].cells[0].text = f"⑬ FTA Agreement: {data['fta']}"
    table.rows[5].cells[1].text = f"⑭ Payment: {data['pay']}"

    item_table = doc.add_table(rows=2, cols=6); item_table.style = 'Table Grid'
    hdr = item_table.rows[0].cells
    for i, txt in enumerate(['Marks', 'Pkgs', 'Description', 'Qty', 'Price', 'Amount']): hdr[i].text = txt
    
    row = item_table.rows[1].cells
    row[0].text, row[1].text, row[2].text = data['marks'], data['pkg_kind'], data['description']
    row[3].text, row[4].text, row[5].text = str(data['qty']), str(data['unit_price']), str(data['amount'])
    
    doc.add_paragraph(f"\nOrigin: MADE IN KOREA (Beneficiary of {data['fta']})")
    doc.add_paragraph(f"Signed by: {data['shipper'].splitlines()[0]}")
    return doc

def create_pl_docx(data):
    """포장명세서(Packing List) .docx 생성"""
    doc = Document()
    doc.add_heading('PACKING LIST', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    table = doc.add_table(rows=4, cols=2); table.style = 'Table Grid'
    table.rows[0].cells[0].text = f"Seller: {data['shipper']}"; table.rows[0].cells[1].text = f"Inv No: {data['inv_no_date']}"
    table.rows[1].cells[0].text = f"Consignee: {data['consignee']}"; table.rows[1].cells[1].text = f"Buyer: {data['buyer']}"
    table.rows[2].cells[0].text = f"Dep. Date: {data['dep_date']}"; table.rows[2].cells[1].text = f"Ref: {data['other_ref']}"
    table.rows[3].cells[0].text = f"Vessel: {data['vessel']} / From: {data['from_port']}"; table.rows[3].cells[1].text = f"To: {data['to_port']}"
    
    item_table = doc.add_table(rows=2, cols=6); item_table.style = 'Table Grid'
    hdr = item_table.rows[0].cells
    for i, txt in enumerate(['Marks', 'Pkgs', 'Goods', 'N.W', 'G.W', 'Meas']): hdr[i].text = txt
    row = item_table.rows[1].cells
    for i, key in enumerate(['marks', 'pkg_kind', 'description', 'net_weight', 'gross_weight', 'measure']): row[i].text = str(data[key])
    return doc

def create_bl_docx(data):
    """선하증권(Bill of Lading) .docx 생성"""
    doc = Document()
    doc.add_heading('BILL OF LADING', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    table = doc.add_table(rows=5, cols=2); table.style = 'Table Grid'
    table.rows[0].cells[0].text = f"Shipper:\n{data['shipper']}"; table.rows[0].cells[1].text = f"B/L No: {data['bl_no']}"
    table.rows[1].cells[0].text = f"Consignee:\n{data['consignee']}"; table.rows[1].cells[1].text = f"Voyage: 1234E"
    table.rows[2].cells[0].text = f"Notify:\nSAME AS CONSIGNEE"; table.rows[2].cells[1].text = f"Vessel: {data['vessel']}"
    table.rows[3].cells[0].text = f"Loading: {data['from_port']}"; table.rows[3].cells[1].text = f"Discharge: {data['to_port']}"
    table.rows[4].cells[0].text = f"Container: ISCU1104"; table.rows[4].cells[1].text = f"Seal No: S123"
    return doc

# --- [Sidebar] 금융 정보 ---
exchange_df = get_exchange_data()
current_rate = exchange_df["환율"].iloc[-1]
delta = round(current_rate - exchange_df["환율"].iloc[-2], 2)

with st.sidebar:
    st.title("💰 금융 & FTA 현황")
    st.metric(label="USD/KRW (2026-01-26)", value=f"{current_rate:,.2f}원", delta=f"{delta}원")
    st.markdown("---")
    st.subheader("📊 실무 가이드")
    st.caption("• 항공(AIR): 해상 대비 운임 약 15% 할증")
    st.caption("• FTA 활용 시 관세 0~5% 수준 절감 가능")
    st.caption("• DDP 조건은 FTA 증빙이 필수입니다.")

# --- [Main Layout] 상단 대시보드 ---
st.title("🚢 Trade Master 2026: FTA & 결제 통합 자동화")

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
st.subheader("📑 거래 상세 및 가격 조건 설정")
with st.form("trade_form"):
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**1. 기본 거래 정보**")
        shipper = st.text_area("수출자(Shipper)", "GILDING TRADING CO., LTD.\nSEOUL, KOREA")
        consignee = st.text_area("수입자(Consignee)", "MONARCH PRO CO., LTD.\nDETROIT, USA")
        from_port = st.text_input("출발지", "BUSAN, KOREA")
        to_port = st.text_input("도착지", "DETROIT, USA")
        vessel = st.text_input("선박/항공편명", "PHEONIC V.123")
    with c2:
        st.markdown("**2. 인코텀즈 및 FTA**")
        incoterms_list = ["EXW", "FCA", "FAS", "FOB", "CPT", "CIP", "CFR", "CIF", "DAP", "DPU", "DDP"]
        selected_term = st.selectbox("Incoterms 2020", incoterms_list)
        fta_list = ["협정 미적용 (기본세율)", "한-미 FTA (KOR-USA)", "한-EU FTA (KOR-EU)", "한-중 FTA (KOR-CHINA)", "한-베트남 FTA (KOR-VIETNAM)", "RCEP"]
        selected_fta = st.selectbox("FTA 협정 선택", fta_list)
        transport_mode = st.radio("운송 수단", ["해상(SEA)", "항공(AIR)"], horizontal=True)
        insurance_type = st.selectbox("적하보험 조건", 
                                    ["선택 안함", "ICC(A) (=ICC(AIR))", "ICC(B) (=ICC(WA))", "ICC(C) (=ICC(FPA))", "ICC(WAIOP)"])
    with c3:
        st.markdown("**3. 품목 및 결제 정보**")
        payment_list = [
            "사전 송금 (Advance Payment)",
            "사후 송금 (Settlement after Shipment)",
            "D/P (Documents against Payment)",
            "D/A (Documents against Acceptance)",
            "일람출급 신용장 (Sight L/C)",
            "기한부 신용장 (Usance L/C)"
        ]
        payment = st.selectbox("결제방식", payment_list)
        description = st.text_input("품명", "NYLON OXFORD")
        qty_input = st.number_input("수량", value=60000)
        unit_price_input = st.number_input("단가(USD)", value=1.00)
        
    st.divider()
    # 인코텀즈 + 결제금융 + FTA 관세가 포함된 실시간 견적 계산
    subtotal = qty_input * unit_price_input
    estimated_total = calculate_estimated_cost(subtotal, selected_term, transport_mode, insurance_type, payment, selected_fta)
    
    st.info(f"💡 **FTA 혜택 및 금융 비용이 반영된 예상 총액:** ${estimated_total:,.2f} (약 {estimated_total * current_rate:,.0f} 원)")

    submitted = st.form_submit_button("🚀 분석 및 서류 생성")

# --- [Bottom] 결과 및 다운로드 ---
if submitted:
    data = {
        "shipper": shipper, "consignee": consignee, "from_port": from_port, "to_port": to_port, "vessel": vessel,
        "inv_no_date": f"INV-2026-{datetime.now().strftime('%m%d')}", "lc_no_date": "LC-2026-001",
        "terms": selected_term, "transport": transport_mode, "insurance": insurance_type, "pay": payment, "fta": selected_fta,
        "description": description, "qty": f"{qty_input:,}", "unit_price": f"{unit_price_input:.2f}",
        "amount": f"{estimated_total:,.2f}", "pkg_kind": "53 C/NO", "net_weight": "1,200 KGS", "gross_weight": "1,208 KGS",
        "marks": "MON/T DETROIT", "measure": "5.8 CBM", "bl_no": "BK-1004", "dep_date": "MAY. 20. 2026",
        "buyer": consignee, "other_ref": "KOREA"
    }
    st.session_state['current_data'] = data

    with st.spinner("AI가 FTA 활용 방안 및 리스크를 분석 중입니다..."):
        risk_prompt = f"""
        당신은 전문 관세사 및 국제금융 전문가입니다. 다음 무역 시나리오를 분석하세요.
        - FTA 협정: {selected_fta}
        - 인코텀즈: {selected_term} / 결제방식: {payment}
        - 운송: {transport_mode} / 보험: {insurance_type}
        - 품목: {description}
        이 협정을 활용하기 위한 원산지 결정 기준(PSR) 충족 필요성과 결제 대금 회수 리스크, 위험 전이 시점을 포함해 상세 대응 전략을 한글로 제시하세요.
        """
        response = client.chat.completions.create(model="gpt-4o", messages=[{"role": "user", "content": risk_prompt}])
        st.session_state['ai_analysis'] = response.choices[0].message.content

if 'ai_analysis' in st.session_state:
    t1, t2 = st.tabs(["💡 AI 전략 가이드", "📥 서류 다운로드"])
    with t1:
        st.markdown(st.session_state['ai_analysis'])
    with t2:
        current_data = st.session_state['current_data']
        doc_files = {
            "Commercial_Invoice_FTA.docx": create_ci_docx(current_data),
            "Packing_List.docx": create_pl_docx(current_data),
            "Bill_of_Lading.docx": create_bl_docx(current_data)
        }
        cols = st.columns(3)
        for i, (name, doc) in enumerate(doc_files.items()):
            bio = BytesIO()
            doc.save(bio)
            cols[i].download_button(label=f"📥 {name}", data=bio.getvalue(), file_name=name,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.success("FTA 정보와 결제 조건이 반영된 모든 서류가 준비되었습니다.")