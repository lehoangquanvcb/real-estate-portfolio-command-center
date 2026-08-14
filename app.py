import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

st.set_page_config(page_title="Trung tâm điều hành dự án bất động sản", page_icon="🏙️", layout="wide", initial_sidebar_state="expanded")


px.defaults.template="plotly_dark"
px.defaults.color_discrete_sequence=["#5AA9E6","#3D7EA6","#64D98B","#F2B84B","#E36B6B","#9B8AFB"]

# ============================= GIAO DIỆN =============================
st.markdown(r"""
<style>
:root{--panel:#0c2234;--line:#20384b;--text:#edf4fb;--muted:#8fa5b8;--green:#20b26b}
[data-testid="stHeader"]{background:#08111c!important;border-bottom:1px solid #172c3d!important;height:3rem!important}[data-testid="stAppViewContainer"]{background:#08111c}
.block-container{padding-top:1.15rem!important;padding-bottom:2rem!important;max-width:100%!important;padding-left:1.15rem!important;padding-right:1.15rem!important}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#071827 0%,#0a2030 100%);border-right:1px solid #244054;min-width:270px!important;max-width:270px!important}
[data-testid="stSidebar"]>div:first-child{padding-top:.45rem!important}[data-testid="stSidebar"] .block-container{padding:.5rem .8rem 1rem!important}
[data-testid="stSidebar"] .stButton{margin:0 0 5px 0!important}[data-testid="stSidebar"] .stButton>button{width:100%!important;height:40px!important;min-height:40px!important;display:flex!important;align-items:center!important;justify-content:flex-start!important;text-align:left!important;padding:0 12px!important;border-radius:6px!important;border:1px solid #284b62!important;background:#0c2a3d!important;color:#eef5fb!important;font-size:.78rem!important;font-weight:650!important;box-shadow:none!important;white-space:nowrap!important}
[data-testid="stSidebar"] .stButton>button:hover{background:#123b56!important;border-color:#3d6f91!important}[data-testid="stSidebar"] .stButton>button[kind="primary"]{background:#123f6b!important;border-color:#4388cf!important;box-shadow:inset 3px 0 0 #64a9f5!important}
[data-testid="stSidebar"] details{border:1px solid #284b62!important;border-radius:7px!important;background:#0a1d2b!important}.section-label{font-size:.66rem;color:#839daf;text-transform:uppercase;letter-spacing:.075em;margin:.78rem 0 .32rem;padding-left:2px}
.topbar{display:flex;align-items:flex-start;border-bottom:1px solid #20384b;padding:5px 0 11px;margin:0 0 8px;position:relative;z-index:3}.brand-title{font-size:1.08rem!important;font-weight:800;line-height:1.3;color:#f4f8fc!important;margin:0;opacity:1!important}.brand-sub{font-size:.76rem;color:#91a8ba;margin-top:5px}.brand-author{color:#58a6ff;font-weight:700}
.update-box{background:#0d2b24;border:1px solid #1a5944;border-radius:7px;padding:8px 11px;color:#62d394;font-size:.72rem;line-height:1.35;margin-top:4px}.update-box b{color:#89e6b4;font-size:.78rem}
.page-title{font-size:1.45rem;font-weight:800;color:#f5f7fa;margin:.72rem 0 .14rem}.page-sub{font-size:.81rem;color:#91a7b9;margin-bottom:.8rem}
.kpi-card{background:linear-gradient(180deg,#0d2436 0%,#0b1d2d 100%);border:1px solid #254156;border-radius:8px;padding:12px 14px;min-height:82px}.kpi-label{font-size:.73rem;color:#9fb0c0}.kpi-value{font-size:1.22rem;font-weight:800;color:#f7fbff;margin-top:3px}.kpi-state{font-size:.68rem;margin-top:4px}.good{color:#43d17d}.warn{color:#f1b94b}.bad{color:#ff6b6b}
[data-testid="stDataFrame"]{border:1px solid #284154;border-radius:7px;overflow:hidden}.stTabs [data-baseweb="tab-list"]{gap:3px;border-bottom:1px solid #20384b;overflow-x:auto}.stTabs [data-baseweb="tab"]{height:38px;padding:0 11px;font-size:.76rem;white-space:nowrap}.stTabs [aria-selected="true"]{background:#123e6a;border-radius:6px 6px 0 0;color:white}
.notice{border:1px solid #254156;background:#0b2030;border-radius:8px;padding:10px 12px;color:#a9bac8;font-size:.78rem}.report-card{border:1px solid #254156;background:linear-gradient(180deg,#0c2132,#091a28);border-radius:9px;padding:13px;min-height:112px}.report-title{font-weight:800;color:#eef5fb}.report-desc{font-size:.76rem;color:#8fa5b8;margin-top:4px}.report-link{font-size:.76rem;color:#56a0ff;margin-top:9px}
@media(max-width:900px){[data-testid="stSidebar"]{min-width:235px!important;max-width:235px!important}.block-container{padding-left:.65rem!important;padding-right:.65rem!important}.page-title{font-size:1.22rem}.brand-title{font-size:.98rem!important}}
</style>
""", unsafe_allow_html=True)

# ============================= DỮ LIỆU =============================
DEFAULT=Path(__file__).with_name("Real_Estate_Project_Master.xlsx")

# Excel đã được Việt hóa; bảng ánh xạ này chỉ để tương thích với các Master cũ hơn nếu người dùng tải lên.
VN_TO_CANON={
"Doanh thu thực tế (tỷ)":"Doanh thu Actual (tỷ)","Thu tiền thực tế (tỷ)":"Thu tiền Actual (tỷ)","CAPEX thực tế (tỷ)":"CAPEX Actual (tỷ)","OPEX thực tế (tỷ)":"OPEX Actual (tỷ)","Lãi vay thực tế (tỷ)":"Lãi vay Actual (tỷ)",
"Số sản phẩm bán":"Sales units","Đạt mốc pháp lý?":"Mốc pháp lý đạt?","Doanh thu ngân sách lũy kế":"Doanh thu Budget YTD","Doanh thu thực tế lũy kế":"Doanh thu Actual YTD","Chênh lệch doanh thu":"Variance DT","Chênh lệch doanh thu (%)":"Variance DT %",
"CAPEX ngân sách lũy kế":"CAPEX Budget YTD","CAPEX thực tế lũy kế":"CAPEX Actual YTD","Chênh lệch CAPEX":"Variance CAPEX","Chênh lệch CAPEX (%)":"Variance CAPEX %","Thu tiền thực tế lũy kế":"Thu tiền Actual YTD","Khoảng thiếu vốn 12T":"Funding gap 12T (tỷ)","Khoảng thiếu vốn 12T (tỷ)":"Funding gap 12T (tỷ)",
"Thay đổi giá bán":"Giá bán Δ","Thay đổi hấp thụ":"Absorption Δ","Thay đổi chi phí":"Chi phí Δ","Thay đổi lãi suất":"Lãi suất Δ","Thay đổi hạn mức tín dụng":"Hạn mức tín dụng Δ","Dự báo doanh thu 24T":"Forecast DT 24T","Dự báo CAPEX 24T":"Forecast CAPEX 24T","Dự báo khoảng thiếu vốn":"Forecast Funding Gap","Dự báo LNST":"Forecast LNST","Kích hoạt dự báo lại":"Trigger reforecast",
"DSCR tối thiểu":"DSCR min","Biên an toàn hiện tại":"Headroom hiện tại","Thời gian dự kiến vi phạm":"Tháng dự kiến breach","Xác suất vi phạm":"Xác suất breach","Lợi nhuận kỳ vọng":"Return proxy","Rủi ro thanh khoản":"Liquidity risk","Rủi ro cam kết tài chính":"Covenant risk","Rủi ro vi phạm":"Xác suất breach","Rủi ro bán hàng":"Sales risk","Điểm chiến lược":"Strategic score","Điểm rủi ro":"Risk score","Điểm lợi nhuận điều chỉnh rủi ro":"RAR score","Nhu cầu vốn 12T":"Vốn yêu cầu 12T",
"Doanh thu ngân sách năm":"Budget DT FY","Dự báo doanh thu còn lại":"Forecast DT còn lại","Doanh thu ước tính mới nhất năm":"LE Doanh thu FY","Chênh lệch DT so ngân sách":"Variance DT vs Budget","Chênh lệch DT (%)":"Variance DT %","CAPEX ngân sách năm":"Budget CAPEX FY","Dự báo CAPEX còn lại":"Forecast CAPEX còn lại","CAPEX ước tính mới nhất năm":"LE CAPEX FY","LNST ước tính":"LNST LE proxy","Trạng thái ước tính":"Trạng thái LE",
"Doanh thu lũy kế":"Doanh thu YTD","Thu tiền lũy kế":"Thu tiền YTD","CAPEX lũy kế":"CAPEX YTD","Phải thu KH ước tính":"Phải thu KH proxy","Tăng HTK/BĐS dở dang lũy kế":"HTK/BĐS dở dang tăng YTD","Phải trả NCC ước tính":"Phải trả NCC proxy","Người mua trả tiền trước ước tính":"Người mua trả tiền trước proxy","Vốn lưu động ròng":"Net Working Capital","Vốn lưu động/Doanh thu":"NWC/Doanh thu",
"Doanh thu ước tính năm":"LE Doanh thu FY","Thuế GTGT hiệu dụng giả định":"VAT effective rate giả định","Thuế TNDN tiền mặt ước tính":"Thuế TNDN cash proxy","Thuế GTGT tiền mặt ước tính":"VAT cash proxy","Tổng thuế/nghĩa vụ tiền mặt":"Tổng cash tax/obligation",
"Tác động giá bán":"Price impact","Tác động hấp thụ/tiến độ":"Absorption/Timing impact","Tác động chậm pháp lý":"Legal delay impact","Phần còn lại do triển khai":"Execution residual","Tác động chi phí":"Cost impact","Tác động lãi vay ước tính":"Interest impact proxy","Cầu nối lợi nhuận ròng ước tính":"Net profit bridge proxy","Yếu tố tác động lớn nhất":"Driver lớn nhất",
"Miền dữ liệu":"Data domain","Chủ dữ liệu":"Data owner","Chủ nghiệp vụ":"Business owner","Quy tắc chốt/khóa":"Cut-off/Lock rule","Trạng thái chất lượng dữ liệu":"Trạng thái DQ","Mã ánh xạ":"Mã mapping","Quy tắc chuyển đổi":"Transform","Kiểm tra hợp lệ":"Validation","Mã lô":"Batch ID","Khóa bản ghi":"Record key","Điểm chất lượng dữ liệu":"DQ Score","Trạng thái lô":"Trạng thái batch",
"Mã nhật ký":"Audit ID","Thời điểm":"Timestamp","Trạng thái rà soát":"Trạng thái review","Mã vai trò":"Role ID","Trạng thái số thực tế":"Actual status","Mã luồng":"Pipeline ID","Đã nhận nguồn":"Source received","Kiểm tra cấu trúc":"Schema check","Kiểm tra ánh xạ":"Mapping check","Kiểm tra chất lượng":"DQ check","Đối chiếu":"Reconciliation","Phê duyệt":"Approval","Kiểm tra khóa kỳ":"Period lock check","Đã chuyển vào dữ liệu chuẩn":"Promoted to Master","Lô gần nhất":"Latest batch","Trạng thái SLA":"SLA status","Hành động tiếp theo":"Next action","Phụ thuộc":"Dependency","Trạng thái kiểm thử":"UAT status","Phê duyệt cuối":"Sign-off","Phụ trách":"Owner","Người rà soát":"Reviewer","Hạn xử lý":"Deadline","Được điều phối tiền?":"Được Cash Pool?","Cam kết tài chính khác":"Covenant khác","Biên hạn mức (tỷ)":"Headroom (tỷ)","Loại khoản vay":"Loại facility"
}

def _clean_loaded_sheet(df):
    if df is None or df.empty:return pd.DataFrame() if df is None else df
    d=df.dropna(how="all").copy()
    d=d.loc[:,~d.columns.astype(str).str.startswith("Unnamed")]
    return d.rename(columns={c:VN_TO_CANON.get(str(c),str(c)) for c in d.columns})

@st.cache_data(show_spinner=False)
def load_book(source):
    x=pd.ExcelFile(source)
    return {name:_clean_loaded_sheet(pd.read_excel(source,sheet_name=name,header=2)) for name in x.sheet_names}

def get_col(df,*names):
    for n in names:
        if n in df.columns:return n
    return None

def num_series(df,*names):
    c=get_col(df,*names)
    return pd.to_numeric(df[c],errors="coerce") if c else pd.Series(0,index=df.index,dtype=float)

def filter_project(df,project):
    if df is None or df.empty or project=="Tất cả dự án":return df
    c=get_col(df,"Mã dự án","Dự án")
    return df[df[c].astype(str)==project].copy() if c else df

def vn_value(v):
    if not isinstance(v,str): return v
    reps=[("GL/BCTC","Sổ cái/BCTC"),("Sales/Collection","Bán hàng/Thu tiền"),("CAPEX/Progress","CAPEX/Tiến độ"),("Cash/Debt","Tiền/Nợ"),("Debt/Covenant","Nợ/Cam kết tài chính"),("Text","Văn bản"),("Number","Số"),("Date","Ngày"),("Boolean","Có/Không"),("Loan register","Sổ theo dõi khoản vay"),("Legal tracker","Theo dõi pháp lý"),("Bank/Treasury","Ngân hàng/Ngân quỹ"),("Revenue/Sales","Doanh thu/Bán hàng"),("AP/Accrual","Phải trả/Chi phí phải trả"),("Intercompany","Giao dịch nội bộ"),("Consolidation","Hợp nhất"),("Tax","Thuế"),("Project Finance","Tài trợ dự án"),("Term Loan","Vay kỳ hạn"),("Acquisition Facility","Khoản vay mua lại"),("Cash sweep","Quét tiền trả nợ"),("Cost overrun test","Kiểm tra vượt chi phí"),("Legal milestone","Mốc pháp lý"),("Legal Gate","Cổng pháp lý"),("upstream cash","chuyển tiền về công ty mẹ"),("cash upstream","chuyển tiền về công ty mẹ"),("facility","hợp đồng tín dụng"),("grace period","ân hạn"),("bridge","khoản vay cầu nối"),("IC loan","vay nội bộ"),("take-out","phương án hoàn trả"),("waiver","miễn trừ"),("refinance","tái tài trợ"),("cash preservation","bảo toàn tiền mặt"),("cash buffer","dự phòng tiền mặt"),("cash-in","dòng tiền vào"),("sales velocity","tốc độ bán hàng"),("milestone","mốc thực hiện"),("equity commitment","cam kết vốn chủ sở hữu"),("posting_date","ngày_hạch_toán"),("entity_code","mã_pháp_nhân"),("project_code","mã_dự_án"),("revenue_amount","doanh_thu"),("paid_capex","capex_đã_thanh_toán"),("ending_balance","số_dư_cuối_kỳ"),("outstanding","dư_nợ"),("GL_YYYYMM_ENTITY.xlsx","SO_CAI_YYYYMM_PHAPNHAN.xlsx"),("SALES_YYYYMM_PROJECT.xlsx","BAN_HANG_YYYYMM_DUAN.xlsx"),("BANK_YYYYMM_ENTITY.xlsx","NGAN_HANG_YYYYMM_PHAPNHAN.xlsx"),("LOAN_REGISTER.xlsx","SO_THEO_DOI_KHOAN_VAY.xlsx"),("LEGAL_TRACKER.xlsx","THEO_DOI_PHAP_LY.xlsx"),("Budget,Commitment,Payment,%HT","Ngân sách,Cam kết,Thanh toán,%HT"),("Debit,Credit,Balance,Ref","Ghi nợ,Ghi có,Số dư,Tham chiếu"),("Limit,Outstanding,Rate,Maturity,Covenant","Hạn mức,Dư nợ,Lãi suất,Đáo hạn,Cam kết tài chính"),("Status,Deadline,Owner,Gate","Trạng thái,Hạn xử lý,Phụ trách,Cổng pháp lý"),("Actual Master","Dữ liệu thực tế chuẩn"),("Actual","Thực tế"),("Budget","Ngân sách"),("Latest Estimate","Ước tính mới nhất"),("Rolling Forecast","Dự báo cuốn chiếu"),("Forecast","Dự báo"),("reforecast","dự báo lại"),("assumption","giả định"),("Owner","Phụ trách"),("Reviewer","Người rà soát"),("Deadline","Hạn xử lý"),("Covenant","Cam kết tài chính"),("Treasury","Ngân quỹ"),("Cash","Tiền"),("Sales","Bán hàng"),("Legal","Pháp lý"),("Execution","Triển khai"),("Cost","Chi phí"),("Price","Giá bán"),("Timing/Other","Tiến độ/Khác"),("Green","Đạt"),("Amber","Cần hoàn thiện"),("NO-GO","CHƯA SẴN SÀNG"),("GO","SẴN SÀNG"),("Source","Nguồn"),("Mapping","Ánh xạ"),("Status","Trạng thái"),("Maturity","Đáo hạn"),("Outstanding","Dư nợ"),("Commitment","Cam kết"),("Payment","Thanh toán"),("Revenue","Doanh thu"),("Bank","Ngân hàng"),("41_Actual_Theo_thang","Số liệu thực tế theo tháng"),("LoanBook Aug-26","Sổ khoản vay tháng 08/2026"),("Project cash","Tiền dự án"),("CRM/ERP","CRM/ERP"),("PMIS/QS","PMIS/QS"),
("Loan register","Sổ theo dõi khoản vay"),("Legal tracker","Theo dõi pháp lý"),
("Limit","Hạn mức"),("Rate","Lãi suất"),("Gate","Cổng pháp lý"),
("Source received","Đã nhận nguồn"),("Schema check","Kiểm tra cấu trúc"),
("Mapping check","Kiểm tra ánh xạ"),("DQ check","Kiểm tra chất lượng"),
("Latest batch","Lô dữ liệu gần nhất"),("SLA status","Trạng thái SLA"),
("Go-Live","Sẵn sàng vận hành"),("UAT","Kiểm thử người dùng"),
("Working Capital","Vốn lưu động"),("Variance","Chênh lệch"),
("Data Pipeline","Luồng dữ liệu"),("Data Governance","Quản trị dữ liệu")]
    out=v
    for a,b in reps: out=out.replace(a,b)
    return out

def display_df(df):
    if df is None:return pd.DataFrame()
    d=df.copy()
    rename={
    "Funding gap (tỷ)":"Khoảng thiếu vốn (tỷ)","Funding gap 12T (tỷ)":"Khoảng thiếu vốn 12T (tỷ)","Budget DT FY":"Doanh thu ngân sách năm","LE Doanh thu FY":"Doanh thu ước tính mới nhất","Variance DT vs Budget":"Chênh lệch DT so ngân sách","LE CAPEX FY":"CAPEX ước tính mới nhất","LNST LE proxy":"LNST ước tính","Trạng thái LE":"Trạng thái ước tính",
    "DSCR min":"DSCR tối thiểu","Headroom hiện tại":"Biên an toàn hiện tại","Tháng dự kiến breach":"Thời gian dự kiến vi phạm","Xác suất breach":"Rủi ro vi phạm","Covenant khác":"Cam kết tài chính khác","Headroom (tỷ)":"Biên hạn mức (tỷ)","Owner":"Phụ trách","Reviewer":"Người rà soát","Deadline":"Hạn xử lý","Driver lớn nhất":"Yếu tố tác động lớn nhất","Net Working Capital":"Vốn lưu động ròng","Data domain":"Miền dữ liệu","DQ Score":"Điểm chất lượng dữ liệu","SLA status":"Trạng thái SLA","Latest batch":"Lô gần nhất","Next action":"Hành động tiếp theo","Source received":"Đã nhận nguồn","Schema check":"Kiểm tra cấu trúc","Mapping check":"Kiểm tra ánh xạ","DQ check":"Kiểm tra chất lượng","Reconciliation":"Đối chiếu","Approval":"Phê duyệt","Period lock check":"Kiểm tra khóa kỳ","Promoted to Master":"Đã chuyển vào dữ liệu chuẩn","UAT status":"Trạng thái kiểm thử","Sign-off":"Phê duyệt cuối","Dependency":"Phụ thuộc","Transform":"Quy tắc chuyển đổi","Validation":"Kiểm tra hợp lệ","Record key":"Khóa bản ghi","Batch ID":"Mã lô","Pipeline ID":"Mã luồng","Audit ID":"Mã nhật ký","Timestamp":"Thời điểm","Actual status":"Trạng thái số thực tế","Data owner":"Chủ dữ liệu","Business owner":"Chủ nghiệp vụ","Cut-off/Lock rule":"Quy tắc chốt/khóa","Mã mapping":"Mã ánh xạ","Loại facility":"Loại khoản vay","Return proxy":"Lợi nhuận kỳ vọng","Liquidity risk":"Rủi ro thanh khoản","Covenant risk":"Rủi ro cam kết tài chính","Sales risk":"Rủi ro bán hàng","Strategic score":"Điểm chiến lược","Risk score":"Điểm rủi ro","RAR score":"Điểm lợi nhuận điều chỉnh rủi ro","Vốn yêu cầu 12T":"Nhu cầu vốn 12T","Trigger reforecast":"Kích hoạt dự báo lại","Forecast DT 24T":"Dự báo doanh thu 24T","Forecast CAPEX 24T":"Dự báo CAPEX 24T","Forecast Funding Gap":"Dự báo khoảng thiếu vốn","Forecast LNST":"Dự báo LNST"}
    d=d.rename(columns=rename)
    for c in d.select_dtypes(include="object").columns:d[c]=d[c].map(vn_value)
    return d

def show_table(df,cols=None,height=None):
    if df is None or df.empty:
        st.info("Chưa có dữ liệu phù hợp.");return
    d=df.copy()
    if cols:d=d[[c for c in cols if c in d.columns]]
    st.dataframe(display_df(d),use_container_width=True,hide_index=True,height=height)


def _num(df, name):
    if df is None or df.empty or name not in df.columns:
        return pd.Series(dtype=float)
    return pd.to_numeric(df[name],errors="coerce")

def safe_bar(df,x,ys,title,ytitle="Tỷ đồng",barmode="group",height=360):
    if df is None or df.empty or x not in df.columns:
        return
    ys=[y for y in ys if y in df.columns]
    if not ys:return
    d=df[[x]+ys].copy()
    for y in ys:d[y]=pd.to_numeric(d[y],errors="coerce").fillna(0)
    fig=px.bar(d,x=x,y=ys,barmode=barmode,title=title)
    _legend_map={
        "Doanh thu Budget YTD":"Doanh thu ngân sách","Doanh thu Actual YTD":"Doanh thu thực tế",
        "CAPEX Budget YTD":"CAPEX ngân sách","CAPEX Actual YTD":"CAPEX thực tế",
        "Budget DT FY":"Doanh thu ngân sách","LE Doanh thu FY":"Doanh thu ước tính mới nhất",
        "Budget CAPEX FY":"CAPEX ngân sách","LE CAPEX FY":"CAPEX ước tính mới nhất",
        "Tổng mức đầu tư (tỷ)":"Tổng mức đầu tư","Doanh thu kỳ vọng (tỷ)":"Doanh thu kỳ vọng",
        "Doanh thu hợp nhất":"Doanh thu hợp nhất","LNST hợp nhất":"LNST hợp nhất",
        "Tổng tài sản hợp nhất":"Tổng tài sản hợp nhất","Tổng nợ hợp nhất":"Tổng nợ hợp nhất",
        "Doanh thu loại trừ":"Doanh thu loại trừ","Chi phí loại trừ":"Chi phí loại trừ",
        "Hạn mức/Khả dụng (tỷ)":"Nguồn vốn khả dụng","Phân bổ đề xuất (tỷ)":"Phân bổ đề xuất",
        "Tiến độ pháp lý":"Tiến độ pháp lý","Tiến độ xây dựng":"Tiến độ xây dựng","Tỷ lệ bán":"Tỷ lệ bán"
    }
    fig.for_each_trace(lambda t:t.update(name=_legend_map.get(t.name,t.name)))
    fig.update_layout(height=height,legend_title_text="Chỉ tiêu",xaxis_title="",yaxis_title=ytitle,margin=dict(l=10,r=10,t=48,b=10))
    st.plotly_chart(fig,use_container_width=True)

def safe_hbar(df,x,y,title,color=None,height=350):
    if df is None or df.empty or x not in df.columns or y not in df.columns:return
    d=df.copy();d[x]=pd.to_numeric(d[x],errors="coerce")
    d=d.dropna(subset=[x,y])
    if d.empty:return
    kwargs={"data_frame":d,"x":x,"y":y,"orientation":"h","title":title}
    if color and color in d.columns:kwargs["color"]=color
    fig=px.bar(**kwargs)
    fig.update_layout(height=height,xaxis_title="",yaxis_title="",margin=dict(l=10,r=10,t=48,b=10))
    st.plotly_chart(fig,use_container_width=True)

def safe_scatter(df,x,y,size=None,color=None,hover=None,title="",height=390):
    if df is None or df.empty or x not in df.columns or y not in df.columns:return
    cols=[x,y]+[c for c in [size,color,hover] if c and c in df.columns]
    d=df[cols].copy()
    d[x]=pd.to_numeric(d[x],errors="coerce");d[y]=pd.to_numeric(d[y],errors="coerce")
    kwargs={"data_frame":d,"x":x,"y":y,"title":title}
    if size and size in d.columns:
        d[size]=pd.to_numeric(d[size],errors="coerce").fillna(0).clip(lower=0)
        if d[size].sum()>0:kwargs["size"]=size
    if color and color in d.columns:kwargs["color"]=color
    if hover and hover in d.columns:kwargs["hover_name"]=hover
    d=d.dropna(subset=[x,y])
    if d.empty:return
    fig=px.scatter(**kwargs)
    fig.update_layout(height=height,margin=dict(l=10,r=10,t=48,b=10))
    st.plotly_chart(fig,use_container_width=True)

def safe_pie(df,names,values,title,height=330):
    if df is None or df.empty or names not in df.columns:return
    d=df.copy()
    if values and values in d.columns:
        d[values]=pd.to_numeric(d[values],errors="coerce").fillna(0)
        fig=px.pie(d,names=names,values=values,title=title,hole=.48)
    else:
        d=d.groupby(names,dropna=False).size().reset_index(name="Số lượng")
        fig=px.pie(d,names=names,values="Số lượng",title=title,hole=.48)
    fig.update_layout(height=height,margin=dict(l=10,r=10,t=48,b=10),legend_title_text="")
    st.plotly_chart(fig,use_container_width=True)

def kpi(label,value,state=None,kind="good"):
    state_html=f'<div class="kpi-state {kind}">{state}</div>' if state else ''
    st.markdown(f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div>{state_html}</div>',unsafe_allow_html=True)

def title(text,sub=""):
    st.markdown(f'<div class="page-title">{text}</div><div class="page-sub">{sub}</div>',unsafe_allow_html=True)

# Sidebar input
with st.sidebar.expander("Dữ liệu nguồn",expanded=False):
    upload=st.file_uploader("Tải file dữ liệu Excel",type=["xlsx"],label_visibility="collapsed")
source=upload if upload else DEFAULT
try:book=load_book(source)
except Exception as e:
    st.error(f"Không đọc được file Master Excel: {e}");st.stop()

def tab(name):return book.get(name,pd.DataFrame()).copy()

portfolio=tab("01_Danh_muc_du_an")
assump=tab("17_Gia_dinh_mo_hinh")
health=tab("16_Suc_khoe_du_an")
project_codes=portfolio["Mã dự án"].dropna().astype(str).tolist() if "Mã dự án" in portfolio.columns else []
def safe_num(x, default=0.0):
    try:
        if pd.isna(x): return default
        return float(x)
    except Exception: return default

def build_project_model(a, scen=None):
    scen=scen or {}
    code=str(a['Mã dự án'])
    months=int(safe_num(a['Số tháng mô hình'],60))
    start=pd.to_datetime(a['Ngày bắt đầu mô hình'])
    delay_days=max(0,safe_num(a['Chậm pháp lý cơ sở (ngày)'])+scen.get('delay_days_add',0))
    delay_m=int(np.ceil(delay_days/30))
    price=1+scen.get('price_change',0)
    absorption=max(0.25,1+scen.get('absorption_change',0))
    cost_mult=1+scen.get('cost_change',0)
    rate=max(0,safe_num(a['Lãi suất vay/năm'])+scen.get('rate_change',0))
    limit=max(0,safe_num(a['Hạn mức vay tối đa (tỷ)'])*(1+scen.get('credit_limit_change',0)))
    total_inv=safe_num(a['Tổng mức đầu tư (tỷ)'])*cost_mult
    total_rev=safe_num(a['Doanh thu kỳ vọng (tỷ)'])*price
    sold=min(0.99,max(0,safe_num(a['Tỷ lệ đã bán hiện tại'])))
    remaining_rev=total_rev*(1-sold)
    gross_margin=safe_num(a['Biên LN gộp'],0.3)
    sell_pct=safe_num(a['CP bán hàng/DT'],0.035)
    admin_pct=safe_num(a['CP QLDN/DT'],0.02)
    tax=safe_num(a['Thuế TNDN'],0.2)
    sales_start=int(safe_num(a['Tháng bắt đầu bán'],1))+delay_m
    sales_period=max(1,int(round(safe_num(a['Số tháng bán'],24)/absorption)))
    recog_start=int(safe_num(a['Tháng bắt đầu ghi nhận DT'],1))+delay_m
    recog_period=max(1,int(round(safe_num(a['Số tháng ghi nhận DT'],18)/absorption)))
    build_period=max(1,int(safe_num(a['Số tháng thi công còn lại'],24)))
    opening_cash=safe_num(a['Tiền đầu kỳ (tỷ)'])
    opening_debt=safe_num(a['Dư nợ đầu kỳ (tỷ)'])
    remaining_invest=total_inv*(1-min(0.95,sold*0.75))
    monthly_capex=remaining_invest/build_period
    monthly_sales=remaining_rev/sales_period
    monthly_recog=remaining_rev/recog_period
    rows=[]
    cash=opening_cash; debt=opening_debt; prev_sales=0.0
    for m in range(1,months+1):
        new_sales=monthly_sales if sales_start <= m < sales_start+sales_period else 0.0
        collection=0.70*new_sales+0.30*prev_sales
        revenue=monthly_recog if recog_start <= m < recog_start+recog_period else 0.0
        cogs=revenue*(1-gross_margin)
        selling=revenue*sell_pct; admin=revenue*admin_pct
        capex=monthly_capex if m<=build_period else 0.0
        interest=debt*rate/12
        pbt=revenue-cogs-selling-admin-interest
        cit=max(0,pbt*tax)
        ni=pbt-cit
        prefin=collection-capex-selling-admin-cit-interest
        headroom=max(0,limit-debt)
        need=max(0,-(cash+prefin))
        draw=min(need,headroom)
        gap=max(0,need-headroom)
        surplus=max(0,cash+prefin+draw)
        repay=min(surplus,debt+draw)
        debt=max(0,debt+draw-repay)
        cash=max(0,cash+prefin+draw-repay)
        cip=max(0, remaining_invest * min(1,m/build_period) - sum(r['Giá vốn (tỷ)'] for r in rows) - cogs)
        rows.append({'Mã dự án':code,'Tháng số':m,'Tháng':start+pd.DateOffset(months=m-1),'Doanh số mới (tỷ)':new_sales,'Thu khách hàng (tỷ)':collection,'Doanh thu (tỷ)':revenue,'Giá vốn (tỷ)':cogs,'CP bán hàng (tỷ)':selling,'CP QLDN (tỷ)':admin,'Chi đầu tư (tỷ)':capex,'Lãi vay (tỷ)':interest,'LNTT (tỷ)':pbt,'Thuế TNDN (tỷ)':cit,'LNST (tỷ)':ni,'Dòng tiền trước tài trợ (tỷ)':prefin,'Vay tăng (tỷ)':draw,'Trả nợ (tỷ)':repay,'Dư nợ cuối kỳ (tỷ)':debt,'Tiền cuối kỳ (tỷ)':cash,'Funding gap (tỷ)':gap,'CIP/Hàng tồn kho proxy (tỷ)':cip})
        prev_sales=new_sales
    return pd.DataFrame(rows)

def annualize(df):
    d=df.copy(); d['Năm mô hình']=((d['Tháng số']-1)//12+1).astype(int)
    flow=['Doanh thu (tỷ)','Giá vốn (tỷ)','CP bán hàng (tỷ)','CP QLDN (tỷ)','Lãi vay (tỷ)','LNTT (tỷ)','Thuế TNDN (tỷ)','LNST (tỷ)','Thu khách hàng (tỷ)','Chi đầu tư (tỷ)','Vay tăng (tỷ)','Trả nợ (tỷ)','Funding gap (tỷ)']
    a=d.groupby('Năm mô hình',as_index=False)[flow].sum()
    end=d.groupby('Năm mô hình',as_index=False).tail(1)[['Năm mô hình','Dư nợ cuối kỳ (tỷ)','Tiền cuối kỳ (tỷ)','CIP/Hàng tồn kho proxy (tỷ)']]
    return a.merge(end,on='Năm mô hình',how='left')


def build_vnfs(m,a):
    d=m.copy(); d['Năm mô hình']=((d['Tháng số']-1)//12+1).astype(int)
    out=[]
    total_inv=safe_num(a['Tổng mức đầu tư (tỷ)'])
    dso=safe_num(a.get('DSO phải thu (ngày)',45),45); dpo=safe_num(a.get('DPO phải trả (ngày)',60),60)
    min_inv=safe_num(a.get('HTK/BĐS dở dang tối thiểu (% TMĐT)',0.08),0.08)
    fixed_open=safe_num(a.get('TSCĐ/BĐSĐT đầu kỳ (tỷ)',100),100); dep_rate=safe_num(a.get('Khấu hao/năm (% TSCĐ)',0.05),0.05)
    other_ca_pct=safe_num(a.get('Thuế GTGT & TSNH khác (% DT)',0.025),0.025); other_pay_pct=safe_num(a.get('Phải trả khác (% DT)',0.04),0.04)
    dividend_pct=safe_num(a.get('Cổ tức/LNST',0),0)
    for y in range(1,6):
        x=d[d['Năm mô hình']==y]
        if x.empty: continue
        revenue=x['Doanh thu (tỷ)'].sum(); cogs=x['Giá vốn (tỷ)'].sum(); sell=x['CP bán hàng (tỷ)'].sum(); admin=x['CP QLDN (tỷ)'].sum()
        interest=x['Lãi vay (tỷ)'].sum(); pbt=x['LNTT (tỷ)'].sum(); tax=x['Thuế TNDN (tỷ)'].sum(); ni=x['LNST (tỷ)'].sum()
        collections=x['Thu khách hàng (tỷ)'].sum(); capex=x['Chi đầu tư (tỷ)'].sum(); draw=x['Vay tăng (tỷ)'].sum(); repay=x['Trả nợ (tỷ)'].sum()
        end=x.iloc[-1]; cash=end['Tiền cuối kỳ (tỷ)']; debt=end['Dư nợ cuối kỳ (tỷ)']
        ar=revenue*dso/365; inventory=max(total_inv*min_inv,total_inv*max(0,1-y/5)*0.50); other_ca=revenue*other_ca_pct
        fixed=max(0,fixed_open+0.08*d.loc[d['Năm mô hình']<=y,'Chi đầu tư (tỷ)'].sum()-fixed_open*dep_rate*y); other_nca=0.03*total_inv
        assets=cash+ar+inventory+other_ca+fixed+other_nca
        ap=cogs*dpo/365; advances=max(0,collections-revenue); taxpay=max(0,tax*0.25); otherpay=revenue*other_pay_pct
        shortdebt=0.35*debt; longdebt=0.65*debt; liabilities=ap+advances+taxpay+otherpay+shortdebt+longdebt
        equity=max(0,assets-liabilities); dividend=max(0,ni)*dividend_pct
        cfo=collections-sell-admin-tax-interest-0.15*capex; cfi=-0.85*capex; cff=draw-repay-dividend
        out.append({'Năm mô hình':y,'Doanh thu':revenue,'Giá vốn':cogs,'Lợi nhuận gộp':revenue-cogs,'Chi phí tài chính':interest,'Chi phí bán hàng':sell,'Chi phí QLDN':admin,'LNTT':pbt,'Thuế TNDN':tax,'LNST':ni,'Tiền':cash,'Phải thu KH':ar,'Hàng tồn kho/BĐS dở dang':inventory,'TSNH khác':other_ca,'TSCĐ & BĐSĐT thuần':fixed,'TSDH khác':other_nca,'Tổng tài sản':assets,'Phải trả người bán':ap,'Người mua trả tiền trước':advances,'Thuế phải nộp':taxpay,'Phải trả khác':otherpay,'Vay ngắn hạn':shortdebt,'Vay dài hạn':longdebt,'Tổng nợ':liabilities,'Vốn CSH':equity,'CFO':cfo,'CFI':cfi,'CFF':cff,'LCT thuần':cfo+cfi+cff})
    return pd.DataFrame(out)

def show_vnfs(fs,title):
    if fs.empty: st.info('Chưa có dữ liệu BCTC.'); return
    years=fs['Năm mô hình'].astype(int).tolist()
    def stmt(rows):
        dat=[]
        for code,label,key in rows:
            dat.append([code,label]+[fs.loc[fs['Năm mô hình']==y,key].iloc[0] if key in fs.columns else 0 for y in years])
        return pd.DataFrame(dat,columns=['Mã số','Chỉ tiêu']+[f'Năm {y}' for y in years])
    st.subheader(title)
    b1,b2,b3,b4=st.tabs(['B01-DN • Bảng cân đối kế toán','B02-DN • Kết quả hoạt động kinh doanh','B03-DN • Lưu chuyển tiền tệ','B04-DN • Thuyết minh'])
    with b1:
        rows=[('110','Tiền và các khoản tương đương tiền','Tiền'),('131','Phải thu ngắn hạn của khách hàng','Phải thu KH'),('140','Hàng tồn kho/BĐS dở dang','Hàng tồn kho/BĐS dở dang'),('150','Tài sản ngắn hạn khác','TSNH khác'),('220','TSCĐ và BĐS đầu tư thuần','TSCĐ & BĐSĐT thuần'),('270','Tài sản dài hạn khác','TSDH khác'),('280','TỔNG CỘNG TÀI SẢN','Tổng tài sản'),('311','Phải trả người bán','Phải trả người bán'),('312','Người mua trả tiền trước','Người mua trả tiền trước'),('313','Thuế phải nộp','Thuế phải nộp'),('319','Phải trả khác','Phải trả khác'),('320','Vay ngắn hạn','Vay ngắn hạn'),('338','Vay dài hạn','Vay dài hạn'),('300','NỢ PHẢI TRẢ','Tổng nợ'),('400','VỐN CHỦ SỞ HỮU','Vốn CSH'),('440','TỔNG CỘNG NGUỒN VỐN','Tổng tài sản')]
        st.dataframe(stmt(rows),use_container_width=True,hide_index=True)
    with b2:
        rows=[('01','Doanh thu bán hàng và cung cấp dịch vụ','Doanh thu'),('11','Giá vốn hàng bán','Giá vốn'),('20','Lợi nhuận gộp','Lợi nhuận gộp'),('22','Chi phí tài chính','Chi phí tài chính'),('25','Chi phí bán hàng','Chi phí bán hàng'),('26','Chi phí quản lý doanh nghiệp','Chi phí QLDN'),('50','Tổng lợi nhuận kế toán trước thuế','LNTT'),('51','Chi phí thuế TNDN','Thuế TNDN'),('60','Lợi nhuận sau thuế TNDN','LNST')]
        st.dataframe(stmt(rows),use_container_width=True,hide_index=True)
    with b3:
        rows=[('20','Lưu chuyển tiền thuần từ HĐKD','CFO'),('30','Lưu chuyển tiền thuần từ HĐ đầu tư','CFI'),('40','Lưu chuyển tiền thuần từ HĐ tài chính','CFF'),('50','Lưu chuyển tiền thuần trong kỳ','LCT thuần'),('70','Tiền cuối kỳ','Tiền')]
        st.dataframe(stmt(rows),use_container_width=True,hide_index=True)
    with b4:
        st.markdown('**Cơ sở lập:** BCTC kế hoạch/pro forma, đơn vị tỷ đồng; cấu trúc theo hệ thống BCTC doanh nghiệp Việt Nam.')
        st.markdown('**Đặc thù BĐS:** cần rà soát hàng tồn kho/BĐS dở dang, người mua trả tiền trước, vốn hóa chi phí đi vay và điều kiện ghi nhận doanh thu theo hồ sơ thực tế.')
        st.info('Từ năm tài chính bắt đầu từ hoặc sau 01/01/2026, áp dụng Thông tư 99/2025/TT-BTC thay Thông tư 200/2014/TT-BTC.')

def all_models(scen=None):
    frames=[]
    for _,a in assump.iterrows(): frames.append(build_project_model(a,scen))
    return pd.concat(frames,ignore_index=True) if frames else pd.DataFrame()



# ============================= ĐIỀU HƯỚNG =============================
if "page" not in st.session_state: st.session_state.page="01. Tổng quan danh mục"
def nav_group(group,items):
    st.sidebar.markdown(f'<div class="section-label">{group}</div>',unsafe_allow_html=True)
    for label in items:
        active=st.session_state.page==label
        if st.sidebar.button(label,key="nav_"+label,use_container_width=True,type="primary" if active else "secondary"):
            st.session_state.page=label; st.rerun()
nav_group("Quản trị & chiến lược",["01. Tổng quan danh mục","02. Bảng điều khiển KPI","03. Bản đồ danh mục","04. Tiến độ tổng thể"])
nav_group("Tài chính & vận hành",["05. BCTC hợp nhất","06. Loại trừ nội bộ","07. Thác nguồn vốn","08. Dòng tiền 60T & sức chịu đựng","09. Sức khỏe & cảnh báo sớm","10. Quyết định TGĐ"])
nav_group("Hiệu suất & dự báo",["11. Hiệu suất & dự báo","12. Phân bổ vốn","13. Bộ báo cáo HĐQT/CFO"])
nav_group("Kế hoạch & kiểm soát",["14. Kế hoạch & kiểm soát","15. Dữ liệu & kiểm soát"])
page=st.session_state.page
st.sidebar.markdown('<div class="section-label">Kịch bản điều hành</div>',unsafe_allow_html=True)
with st.sidebar.expander("Điều chỉnh giả định",expanded=False):
    price_change=st.slider("Giá bán",-30,20,0,1)/100
    absorption_change=st.slider("Tốc độ hấp thụ",-50,30,0,5)/100
    cost_change=st.slider("Chi phí đầu tư/xây dựng",-10,30,0,1)/100
    rate_change=st.slider("Lãi suất vay",-2.0,5.0,0.0,.25)/100
    delay_add=st.slider("Chậm pháp lý bổ sung (ngày)",0,365,0,15)
    credit_limit_change=st.slider("Thay đổi hạn mức tín dụng",-50,30,0,5)/100
scen={"price_change":price_change,"absorption_change":absorption_change,"cost_change":cost_change,"rate_change":rate_change,"delay_days_add":delay_add,"credit_limit_change":credit_limit_change}
model_all=all_models(scen)

st.sidebar.markdown('---')
st.sidebar.markdown('<div class="report-card"><div class="report-title">Hướng dẫn sử dụng</div><div class="report-desc">Chọn trang điều hành, chọn dự án ở đầu màn hình và tập trung xử lý các ngoại lệ màu vàng/đỏ.</div><div class="report-link">ⓘ Số thực tế đã khóa không bị dự báo ghi đè.</div></div>',unsafe_allow_html=True)

# Header + bộ lọc dự án
left,right=st.columns([5.1,1.5],vertical_alignment="top")
with left:
    st.markdown('<div class="topbar"><div><div class="brand-title">🏙️ Trung tâm điều hành danh mục, phát triển & đầu tư dự án bất động sản</div><div class="brand-sub">Nền tảng quản trị đa dự án dành cho doanh nghiệp trong nước &nbsp; | &nbsp; Tác giả: <span class="brand-author">Lê Hoàng Quân</span></div></div></div>',unsafe_allow_html=True)
with right:
    st.caption("Chọn dự án")
    choices=["Tất cả dự án"]+project_codes
    selected=st.selectbox("Chọn dự án",choices,label_visibility="collapsed")
    now=datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).strftime("%d/%m/%Y %H:%M")
    st.markdown(f'<div class="update-box">◷ Cập nhật gần nhất<br><b>{now}</b></div>',unsafe_allow_html=True)

# Data tables
entities=tab("29_Danh_muc_phap_nhan");treasury=tab("32_Treasury_Cash_Pool");debtbook=tab("33_No_vay_Covenant")
elimin=tab("31_Loai_tru_hop_nhat");waterfall=tab("34_Waterfall_Nguon_von");consdata=tab("35_BCTC_Hop_nhat_Data")
ceo=tab("39_Quyet_dinh_TGD")
actual=tab("41_Actual_Theo_thang");bva=tab("42_Budget_vs_Actual");rolling=tab("43_Rolling_Forecast");covfc=tab("44_Covenant_Forecast");alloc=tab("45_Phan_bo_von")
versions=tab("47_Quan_ly_Phien_ban");workflow=tab("48_Workflow_Phe_duyet");close=tab("49_Monthly_Close");latest=tab("50_Latest_Estimate");wc=tab("51_Working_Capital");tax=tab("52_Ke_hoach_Thue");bridge=tab("53_Variance_Bridge");dq=tab("54_Quan_tri_Du_lieu")
source_cfg=tab("56_Cau_hinh_Import");mapping=tab("57_Tu_dien_Mapping");staging=tab("58_Staging_Import");recon=tab("59_Doi_chieu_Du_lieu");audit=tab("60_Audit_Trail");roles=tab("61_Phan_quyen");locks=tab("62_Khoa_Ky");pipe=tab("63_Data_Pipeline");ready=tab("64_GoLive_Readiness");ops=tab("65_Operational_Control")

# Lọc theo dự án ở cấp trang
def pf(df):return filter_project(df,selected)

# ============================= CÁC TRANG =============================
if page.startswith("01."):
    title("Tổng quan danh mục","Tổng hợp nhanh quy mô, nguồn vốn, sức khỏe và các quyết định trọng yếu của toàn danh mục.")
    p=portfolio if selected=="Tất cả dự án" else pf(portfolio)
    total_inv=num_series(p,"Tổng mức đầu tư (tỷ)").sum();total_rev=num_series(p,"Doanh thu kỳ vọng (tỷ)").sum()
    gap=num_series(pf(treasury),"Funding gap 12T (tỷ)","Khoảng thiếu vốn 12T (tỷ)").sum();debt=num_series(pf(debtbook),"Dư nợ hiện tại (tỷ)").sum()
    ks=st.columns(4)
    with ks[0]:kpi("Tổng mức đầu tư",f"{total_inv:,.0f} tỷ")
    with ks[1]:kpi("Doanh thu kỳ vọng",f"{total_rev:,.0f} tỷ")
    with ks[2]:kpi("Khoảng thiếu vốn 12T",f"{gap:,.1f} tỷ", "Cần xử lý" if gap>0 else "Trong ngưỡng","bad" if gap>0 else "good")
    with ks[3]:kpi("Dư nợ hiện tại",f"{debt:,.1f} tỷ")
    st.markdown("#### Danh mục dự án")
    show_table(p,["Mã dự án","Tên dự án","Địa phương","Giai đoạn","Tổng mức đầu tư (tỷ)","Doanh thu kỳ vọng (tỷ)","Vấn đề chính"],330)
    c1,c2=st.columns([1.35,1])
    with c1:
        safe_bar(p,"Mã dự án",["Tổng mức đầu tư (tỷ)","Doanh thu kỳ vọng (tỷ)"],"Quy mô đầu tư & doanh thu kỳ vọng")
    with c2:
        h=pf(health)
        safe_pie(h,"Xếp loại",None,"Cơ cấu sức khỏe danh mục")
    c1,c2=st.columns([1.1,1])
    with c1:
        h=pf(health);show_table(h,["Mã dự án","Tên dự án","Điểm sức khỏe","Xếp loại","Mức cảnh báo"],260)
    with c2:
        d=pf(ceo);show_table(d,["Ưu tiên","Pháp nhân/Dự án","Vấn đề","Mức độ","Khuyến nghị","Chủ trì","Hạn xử lý"],260)

elif page.startswith("02."):
    title("Bảng điều khiển KPI","Theo dõi ngoại lệ về doanh thu, CAPEX, vốn, nợ và sức khỏe dự án.")
    d=pf(bva)
    warn=(d["Cảnh báo"].astype(str)=="CẢNH BÁO").sum() if "Cảnh báo" in d.columns else 0
    gap=num_series(d,"Funding gap 12T (tỷ)","Khoảng thiếu vốn 12T (tỷ)").sum();rev=num_series(d,"Doanh thu Actual YTD","Doanh thu thực tế lũy kế").sum();capex=num_series(d,"CAPEX Actual YTD","CAPEX thực tế lũy kế").sum()
    c=st.columns(4)
    with c[0]:kpi("Dự án cảnh báo",str(int(warn)),"Ngoại lệ cần xử lý","bad" if warn else "good")
    with c[1]:kpi("Doanh thu thực tế lũy kế",f"{rev:,.1f} tỷ")
    with c[2]:kpi("CAPEX thực tế lũy kế",f"{capex:,.1f} tỷ")
    with c[3]:kpi("Khoảng thiếu vốn 12T",f"{gap:,.1f} tỷ","Cảnh báo" if gap>0 else "Ổn định","bad" if gap>0 else "good")
    show_table(d,height=360)
    if not d.empty:
        fig=px.bar(d,x="Mã dự án",y=[c for c in ["Doanh thu Budget YTD","Doanh thu Actual YTD"] if c in d.columns],barmode="group")
        fig.update_layout(title="Ngân sách và thực tế – Doanh thu",legend_title_text="Chỉ tiêu",yaxis_title="Tỷ đồng",xaxis_title="Dự án")
        fig.for_each_trace(lambda t:t.update(name={"Doanh thu Budget YTD":"Ngân sách","Doanh thu Actual YTD":"Thực tế"}.get(t.name,t.name)))
        st.plotly_chart(fig,use_container_width=True)
    c1,c2=st.columns(2)
    with c1:
        safe_bar(d,"Mã dự án",[c for c in ["CAPEX Budget YTD","CAPEX Actual YTD"] if c in d.columns],"Ngân sách và thực tế – CAPEX")
    with c2:
        gap_col=get_col(d,"Funding gap 12T (tỷ)","Khoảng thiếu vốn 12T (tỷ)")
        if gap_col:safe_hbar(d,gap_col,"Mã dự án","Khoảng thiếu vốn 12 tháng")

elif page.startswith("03."):
    title("Bản đồ danh mục","Nhìn danh mục theo giai đoạn, quy mô đầu tư, sức khỏe và khoảng thiếu vốn.")
    p=portfolio.copy()
    # Portfolio đã có Điểm sức khỏe; chỉ bổ sung Xếp loại từ sheet sức khỏe để tránh sinh _x/_y.
    if not health.empty and "Mã dự án" in health.columns:
        extra=[c for c in ["Mã dự án","Xếp loại"] if c in health.columns]
        if len(extra)>1:p=p.merge(health[extra].drop_duplicates("Mã dự án"),on="Mã dự án",how="left")
    if not treasury.empty and "Mã dự án" in treasury.columns:
        gap_col=get_col(treasury,"Funding gap 12T (tỷ)","Khoảng thiếu vốn 12T (tỷ)")
        if gap_col:
            tg=treasury[["Mã dự án",gap_col]].copy().rename(columns={gap_col:"Khoảng thiếu vốn 12T (tỷ)"})
            p=p.merge(tg.drop_duplicates("Mã dự án"),on="Mã dự án",how="left")
    p=pf(p)
    show_table(p,["Mã dự án","Tên dự án","Địa phương","Giai đoạn","Tổng mức đầu tư (tỷ)","Doanh thu kỳ vọng (tỷ)","Điểm sức khỏe","Xếp loại","Khoảng thiếu vốn 12T (tỷ)"],300)
    c1,c2=st.columns([1.45,1])
    with c1:
        safe_scatter(p,"Tổng mức đầu tư (tỷ)","Điểm sức khỏe","Doanh thu kỳ vọng (tỷ)","Giai đoạn","Tên dự án","Quy mô đầu tư – Điểm sức khỏe")
    with c2:
        safe_pie(p,"Giai đoạn","Tổng mức đầu tư (tỷ)","Cơ cấu tổng mức đầu tư theo giai đoạn")
    gap_col=get_col(p,"Khoảng thiếu vốn 12T (tỷ)","Funding gap 12T (tỷ)")
    if gap_col:
        safe_hbar(p,gap_col,"Mã dự án","Khoảng thiếu vốn theo dự án")

elif page.startswith("04."):
    title("Tiến độ tổng thể","Tập trung vào đường găng pháp lý, tiến độ thực hiện và các mốc đang quá hạn.")
    legal=pf(tab("14_Duong_gang_PL"));actions=pf(tab("07_Hanh_dong"))
    p=pf(portfolio)
    c1,c2=st.columns(2)
    with c1:
        safe_bar(p,"Mã dự án",[c for c in ["Tiến độ pháp lý","Tiến độ xây dựng","Tỷ lệ bán"] if c in p.columns],"Tiến độ pháp lý – xây dựng – bán hàng",ytitle="Tỷ lệ")
    with c2:
        if not actions.empty:
            status_col=get_col(actions,"Trạng thái","Mức độ")
            if status_col:safe_pie(actions,status_col,None,"Cơ cấu công việc theo trạng thái")
    t1,t2=st.tabs(["Đường găng pháp lý","Danh sách hành động"])
    with t1:show_table(legal,height=360)
    with t2:show_table(actions,height=360)

elif page.startswith("05."):
    title("Báo cáo tài chính hợp nhất","B01-DN, B02-DN và B03-DN hợp nhất phục vụ quản trị.")
    if selected!="Tất cả dự án" and not assump.empty:
        a=assump[assump["Mã dự án"].astype(str)==selected]
        if not a.empty: show_vnfs(build_vnfs(build_project_model(a.iloc[0],scen),a.iloc[0]),f"BCTC kế hoạch – {selected}")
    else:
        if not consdata.empty:
            c1,c2=st.columns(2)
            with c1:safe_bar(consdata,"Năm mô hình",[c for c in ["Doanh thu hợp nhất","LNST hợp nhất"] if c in consdata.columns],"Doanh thu & LNST hợp nhất")
            with c2:safe_bar(consdata,"Năm mô hình",[c for c in ["Tổng tài sản hợp nhất","Tổng nợ hợp nhất"] if c in consdata.columns],"Tài sản & nợ hợp nhất")
        t1,t2,t3=st.tabs(["Bảng cân đối kế toán","Kết quả hoạt động kinh doanh","Lưu chuyển tiền tệ"])
        with t1:show_table(tab("36_B01_DN_Hop_nhat"),height=500)
        with t2:show_table(tab("37_B02_DN_Hop_nhat"),height=500)
        with t3:show_table(tab("38_B03_DN_Hop_nhat"),height=500)

elif page.startswith("06."):
    title("Loại trừ nội bộ","Kiểm soát giao dịch nội bộ và bút toán loại trừ trước khi hợp nhất.")
    ic=tab("30_Giao_dich_noi_bo")
    c1,c2=st.columns(2)
    with c1:
        safe_bar(elimin,"Năm mô hình",[c for c in ["Doanh thu loại trừ","Chi phí loại trừ"] if c in elimin.columns],"Doanh thu & chi phí nội bộ loại trừ")
    with c2:
        if not ic.empty and "Loại giao dịch" in ic.columns:safe_pie(ic,"Loại giao dịch","Giá trị phát sinh (tỷ)" if "Giá trị phát sinh (tỷ)" in ic.columns else None,"Cơ cấu giao dịch nội bộ")
    t1,t2=st.tabs(["Giao dịch nội bộ","Bút toán loại trừ"])
    with t1:show_table(ic,height=360)
    with t2:show_table(elimin,height=360)

elif page.startswith("07."):
    title("Thác nguồn vốn","Ưu tiên nguồn vốn theo điều kiện pháp lý, chi phí vốn và khả năng sử dụng.")
    d=pf(waterfall)
    if not d.empty:
        safe_bar(d,"Mã dự án",[c for c in ["Hạn mức/Khả dụng (tỷ)","Phân bổ đề xuất (tỷ)"] if c in d.columns],"Nguồn vốn khả dụng & phân bổ đề xuất",barmode="group")
    show_table(d,height=420)

elif page.startswith("08."):
    title("Dòng tiền 60 tháng & sức chịu đựng","Đánh giá thu tiền, chi đầu tư, dư nợ và khoảng thiếu vốn theo kịch bản.")
    m=model_all if selected=="Tất cả dự án" else model_all[model_all["Mã dự án"].astype(str)==selected]
    agg=m.groupby("Tháng",as_index=False)[["Thu khách hàng (tỷ)","Chi đầu tư (tỷ)","Dư nợ cuối kỳ (tỷ)","Funding gap (tỷ)"]].sum() if not m.empty else pd.DataFrame()
    if not agg.empty:
        fig=go.Figure();fig.add_trace(go.Scatter(x=agg["Tháng"],y=agg["Thu khách hàng (tỷ)"],name="Thu khách hàng"));fig.add_trace(go.Scatter(x=agg["Tháng"],y=agg["Chi đầu tư (tỷ)"],name="Chi đầu tư"));fig.add_trace(go.Scatter(x=agg["Tháng"],y=agg["Dư nợ cuối kỳ (tỷ)"],name="Dư nợ"));fig.update_layout(hovermode="x unified",legend_title_text="Chỉ tiêu",yaxis_title="Tỷ đồng",xaxis_title="Tháng")
        st.plotly_chart(fig,use_container_width=True)
        c=st.columns(3);peak=agg["Dư nợ cuối kỳ (tỷ)"].max();gap=agg["Funding gap (tỷ)"].max();cash_in=agg["Thu khách hàng (tỷ)"].sum()
        with c[0]:kpi("Dư nợ đỉnh",f"{peak:,.1f} tỷ")
        with c[1]:kpi("Khoảng thiếu vốn đỉnh",f"{gap:,.1f} tỷ","Cảnh báo" if gap>0 else "Không thiếu vốn","bad" if gap>0 else "good")
        with c[2]:kpi("Tổng thu khách hàng",f"{cash_in:,.1f} tỷ")

elif page.startswith("09."):
    title("Sức khỏe dự án & cảnh báo sớm","Theo dõi điểm sức khỏe, rủi ro và tín hiệu cần can thiệp.")
    h=pf(health);show_table(h,height=360)
    if not h.empty and "Điểm sức khỏe" in h.columns:
        fig=px.bar(h.sort_values("Điểm sức khỏe"),x="Điểm sức khỏe",y="Tên dự án",orientation="h",color="Xếp loại",title="Điểm sức khỏe dự án")
        fig.update_layout(legend_title_text="Xếp loại");st.plotly_chart(fig,use_container_width=True)

elif page.startswith("10."):
    title("Quyết định TGĐ","Danh sách quyết định cần xử lý, chủ trì và thời hạn đóng việc.")
    d=pf(ceo)
    c1,c2=st.columns([1,1.4])
    with c1:
        if not d.empty and "Mức độ" in d.columns:safe_pie(d,"Mức độ",None,"Cơ cấu quyết định theo mức độ")
    with c2:
        if not d.empty and "Chủ trì" in d.columns:
            tmp=d.groupby("Chủ trì",dropna=False).size().reset_index(name="Số việc")
            safe_hbar(tmp,"Số việc","Chủ trì","Khối lượng quyết định theo đầu mối")
    show_table(d,height=400)

elif page.startswith("11."):
    title("Hiệu suất & dự báo","Ngân sách so với thực tế, dự báo cuốn chiếu và ước tính mới nhất.")
    d=pf(latest)
    c1,c2=st.columns(2)
    with c1:safe_bar(d,"Mã dự án",[c for c in ["Budget DT FY","LE Doanh thu FY"] if c in d.columns],"Doanh thu ngân sách & ước tính mới nhất")
    with c2:safe_bar(d,"Mã dự án",[c for c in ["Budget CAPEX FY","LE CAPEX FY"] if c in d.columns],"CAPEX ngân sách & ước tính mới nhất")
    t1,t2,t3=st.tabs(["Ngân sách so với thực tế","Dự báo cuốn chiếu","Ước tính mới nhất"])
    with t1:show_table(pf(bva),height=360)
    with t2:show_table(pf(rolling),height=360)
    with t3:show_table(d,height=360)

elif page.startswith("12."):
    title("Phân bổ vốn","Xếp hạng ưu tiên vốn dựa trên lợi nhuận kỳ vọng, rủi ro và nhu cầu vốn.")
    d=pf(alloc)
    c1,c2=st.columns(2)
    with c1:safe_hbar(d,"RAR score","Dự án","Điểm lợi nhuận điều chỉnh rủi ro","Quyết định")
    with c2:
        gap_col=get_col(d,"Funding gap","Khoảng thiếu vốn")
        if gap_col:safe_hbar(d,gap_col,"Dự án","Khoảng thiếu vốn & ưu tiên phân bổ")
    show_table(d,height=400)

elif page.startswith("13."):
    title("Bộ báo cáo HĐQT/CFO – 3–5 phút","Tóm tắt các vấn đề trọng yếu cần HĐQT/CFO quyết định và chỉ đạo trong kỳ.")
    d=pf(latest)
    gap=num_series(d,"Funding gap 12T (tỷ)","Khoảng thiếu vốn 12T (tỷ)").sum();rev=num_series(d,"LE Doanh thu FY","Doanh thu ước tính mới nhất năm").sum();var=num_series(d,"Variance DT vs Budget","Chênh lệch DT so ngân sách").sum();capex=num_series(d,"LE CAPEX FY","CAPEX ước tính mới nhất năm").sum()
    overdue=(close["Trạng thái"].astype(str)=="Quá hạn").sum() if "Trạng thái" in close.columns else 0;highcov=(covfc["Xác suất breach"].astype(str)=="Cao").sum() if "Xác suất breach" in covfc.columns else 0
    ks=st.columns(6)
    vals=[("Doanh thu ước tính",f"{rev:,.1f} tỷ","Ổn định","good"),("Chênh lệch DT",f"{var:,.1f} tỷ","Cảnh báo" if var<0 else "Tích cực","bad" if var<0 else "good"),("CAPEX ước tính",f"{capex:,.1f} tỷ","Theo dõi","warn"),("Khoảng thiếu vốn",f"{gap:,.1f} tỷ","Cảnh báo" if gap>0 else "Ổn định","bad" if gap>0 else "good"),("Khóa sổ quá hạn",str(int(overdue)),"Cần xử lý" if overdue else "Đạt","bad" if overdue else "good"),("Khoản vay rủi ro cao",str(int(highcov)),"Cần xử lý" if highcov else "Đạt","bad" if highcov else "good")]
    for col,(lab,val,state,kind) in zip(ks,vals):
        with col:kpi(lab,val,state,kind)
    c1,c2=st.columns(2)
    with c1:safe_bar(d,"Mã dự án",[c for c in ["Budget DT FY","LE Doanh thu FY"] if c in d.columns],"Doanh thu: Ngân sách & Ước tính mới nhất")
    with c2:
        gap_col=get_col(d,"Funding gap 12T (tỷ)","Khoảng thiếu vốn 12T (tỷ)")
        if gap_col:safe_hbar(d,gap_col,"Mã dự án","Khoảng thiếu vốn 12 tháng")
    f1,f2,f3,f4=st.columns([1,1,1,.72]);
    with f1:status_filter=st.selectbox("Lọc theo trạng thái",["Tất cả","CẦN HÀNH ĐỘNG","TRONG NGƯỠNG"])
    with f2:project_filter=st.selectbox("Lọc theo dự án",["Tất cả"]+project_codes)
    with f3:topic=st.selectbox("Lọc theo chủ đề",["Tất cả","Doanh thu","CAPEX","Nguồn vốn","Cam kết tài chính"])
    with f4:
        st.caption("Xuất dữ liệu")
        st.download_button("📗 Xuất Excel",data=DEFAULT.read_bytes(),file_name="Real_Estate_Project_Master.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    table=d.copy();
    if status_filter!="Tất cả" and "Trạng thái LE" in table.columns:table=table[table["Trạng thái LE"].astype(str)==status_filter]
    if project_filter!="Tất cả" and "Mã dự án" in table.columns:table=table[table["Mã dự án"].astype(str)==project_filter]
    show_table(table,["Mã dự án","Budget DT FY","LE Doanh thu FY","Variance DT vs Budget","LE CAPEX FY","Funding gap 12T (tỷ)","Trạng thái LE","Hành động"],330)
    st.markdown('<div class="notice">ⓘ <b>Ghi chú:</b> Ước tính mới nhất là số điều hành; Ngân sách là mốc so sánh; số Thực tế đã khóa không bị ghi đè.</div>',unsafe_allow_html=True)
    st.markdown("#### Quyết định ưu tiên")
    show_table(pf(ceo),["Ưu tiên","Pháp nhân/Dự án","Vấn đề","Mức độ","Khuyến nghị","Chủ trì","Hạn xử lý","Trạng thái"],260)
    st.markdown("#### Bộ báo cáo nhanh")
    cards=st.columns(4)
    for c,(t,dsc) in zip(cards,[("Báo cáo HĐQT","Các vấn đề trọng yếu cần HĐQT quyết định."),("Báo cáo CFO","Tài chính, vốn, dòng tiền và cam kết tài chính."),("Báo cáo TGĐ","Tiến độ triển khai, kế hoạch hành động và KPI."),("Hướng dẫn & giải thích","Giải thích KPI, công thức và cách sử dụng.")]):
        with c:st.markdown(f'<div class="report-card"><div class="report-title">{t}</div><div class="report-desc">{dsc}</div><div class="report-link">Xem nội dung →</div></div>',unsafe_allow_html=True)

elif page.startswith("14."):
    title("Kế hoạch & kiểm soát","Quản lý phiên bản, luồng phê duyệt, khóa sổ, vốn lưu động và nghĩa vụ tài chính.")
    c1,c2,c3=st.columns(3)
    with c1:
        if not close.empty and "Trạng thái" in close.columns:safe_pie(close,"Trạng thái",None,"Tình trạng khóa sổ")
    with c2:
        d=pf(wc); safe_hbar(d,"Net Working Capital","Mã dự án","Vốn lưu động ròng") if "Net Working Capital" in d.columns else None
    with c3:
        d=pf(tax); safe_hbar(d,"Đến hạn 90 ngày","Mã dự án","Nghĩa vụ đến hạn 90 ngày") if "Đến hạn 90 ngày" in d.columns else None
    t1,t2,t3,t4=st.tabs(["Phiên bản & phê duyệt","Khóa sổ tháng","Vốn lưu động","Thuế & nghĩa vụ"])
    with t1:show_table(versions,height=300);show_table(workflow,height=300)
    with t2:show_table(close,height=480)
    with t3:show_table(pf(wc),height=440)
    with t4:show_table(pf(tax),height=440)

else:
    title("Dữ liệu & kiểm soát","Kiểm soát nhập dữ liệu, ánh xạ, đối chiếu, nhật ký, phân quyền, khóa kỳ và sẵn sàng vận hành.")
    c1,c2,c3=st.columns(3)
    with c1:
        if not staging.empty:safe_hbar(staging,"DQ Score","Batch ID","Điểm chất lượng lô dữ liệu")
    with c2:
        if not pipe.empty and "SLA status" in pipe.columns:safe_pie(pipe,"SLA status",None,"Trạng thái SLA luồng dữ liệu")
    with c3:
        if not ready.empty:safe_hbar(ready,"Điểm hiện tại","Hạng mục","Mức sẵn sàng vận hành","Trạng thái")
    t1,t2,t3,t4,t5=st.tabs(["Nguồn & ánh xạ","Dữ liệu chờ & đối chiếu","Nhật ký & phân quyền","Khóa kỳ & luồng dữ liệu","Sẵn sàng vận hành"])
    with t1:show_table(source_cfg,height=270);show_table(mapping,height=300)
    with t2:show_table(staging,height=270);show_table(recon,height=270)
    with t3:show_table(audit,height=270);show_table(roles,height=300)
    with t4:show_table(locks,height=270);show_table(pipe,height=300)
    with t5:show_table(ready,height=340);show_table(ops,height=260)
