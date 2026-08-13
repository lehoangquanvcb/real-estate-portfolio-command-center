import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

st.set_page_config(page_title="Trung tâm điều hành dự án bất động sản", page_icon="🏙️", layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
.block-container{padding-top:1.15rem!important;padding-bottom:2rem;max-width:100%!important}[data-testid="stSidebar"]{background:linear-gradient(180deg,#0b1724,#0e1d2b);border-right:1px solid #26394b}.re-header{border-bottom:1px solid #26394b;padding:2px 0 14px;margin-bottom:10px}.re-title{font-size:clamp(1.35rem,2.05vw,2.05rem);font-weight:800;line-height:1.18;margin:0;color:#f4f7fb;max-width:1150px}.re-sub{font-size:.88rem;color:#9fb0c0;margin-top:8px}.re-author{color:#5aa2ff;font-weight:700}div[data-testid="stMetric"]{background:rgba(20,42,62,.72);border:1px solid #26394b;border-radius:10px;padding:10px 12px}div[data-testid="stMetricValue"]{font-size:1.25rem}[data-testid="stDataFrame"]{border:1px solid #26394b;border-radius:9px;overflow:hidden}.stTabs [data-baseweb="tab-list"]{gap:2px;border-bottom:1px solid #26394b;overflow-x:auto}.stTabs [data-baseweb="tab"]{height:42px;padding:0 12px;font-size:.80rem;white-space:nowrap}.stTabs [aria-selected="true"]{background:#123a67;border-radius:7px 7px 0 0;color:#fff}@media(max-width:700px){.block-container{padding-left:.65rem!important;padding-right:.65rem!important}.re-title{font-size:1.25rem}.re-sub{font-size:.76rem}}
</style>""",unsafe_allow_html=True)
st.markdown("""<div class="re-header"><div class="re-title">🏙️ Trung tâm điều hành danh mục, phát triển &amp; đầu tư dự án bất động sản</div><div class="re-sub">Nền tảng quản trị đa dự án dành cho doanh nghiệp trong nước &nbsp;|&nbsp; Tác giả: <span class="re-author">Lê Hoàng Quân</span> &nbsp;|&nbsp; Phiên bản V10</div></div>""",unsafe_allow_html=True)

VI_COL={'Funding gap (tỷ)':'Khoảng thiếu vốn (tỷ)','Khoảng thiếu vốn 12T (tỷ)':'Khoảng thiếu vốn 12T (tỷ)','Budget DT FY':'Ngân sách doanh thu năm','LE DT FY':'Ước tính mới nhất doanh thu năm','LE Doanh thu FY':'Ước tính mới nhất doanh thu năm','Var DT':'Chênh lệch doanh thu','LE CAPEX':'CAPEX ước tính mới nhất','Driver lớn nhất':'Yếu tố tác động lớn nhất','Action':'Hành động','Owner':'Phụ trách','Deadline':'Hạn xử lý','Data domain':'Miền dữ liệu','DQ Score':'Điểm chất lượng dữ liệu','SLA status':'Trạng thái SLA','Latest batch':'Lô dữ liệu gần nhất','Next action':'Hành động tiếp theo','Source received':'Đã nhận nguồn','Schema check':'Kiểm tra cấu trúc','Ánh xạ check':'Kiểm tra ánh xạ','DQ check':'Kiểm tra chất lượng','Đối chiếu dữ liệu':'Đối chiếu','Approval':'Phê duyệt','Period lock check':'Kiểm tra khóa kỳ','Promoted to Master':'Đã chuyển vào Master','Audit required?':'Yêu cầu nhật ký?','Actual status':'Trạng thái số thực tế','UAT status':'Trạng thái kiểm thử','Sign-off':'Phê duyệt cuối','Dependency':'Phụ thuộc','Readiness Score':'Điểm sẵn sàng','Go-Live Decision':'Quyết định vận hành','Net Vốn lưu động':'Vốn lưu động ròng','Net profit bridge proxy':'Cầu nối lợi nhuận ròng','Price impact':'Tác động giá bán','Absorption/Timing impact':'Tác động hấp thụ/tiến độ','Legal delay impact':'Tác động chậm pháp lý','Execution residual':'Phần còn lại do triển khai','Cost impact':'Tác động chi phí','Interest impact proxy':'Tác động lãi vay','Reviewer':'Người rà soát'}
VI_VAL={'Execution':'Triển khai','Cost':'Chi phí','Legal':'Pháp lý','Price':'Giá bán','Timing/Other':'Tiến độ/Khác','Green':'Đạt','Amber':'Cần hoàn thiện','NO-GO':'CHƯA VẬN HÀNH','GO':'SẴN SÀNG VẬN HÀNH','Promoted':'Đã chuyển vào Master','Approval':'Phê duyệt','DQ check':'Kiểm tra chất lượng'}
_orig_dataframe=st.dataframe
def _vi_df(data):
    if not isinstance(data,pd.DataFrame): return data
    d=data.copy().rename(columns=VI_COL)
    for c in d.columns:
        if d[c].dtype=='object': d[c]=d[c].replace(VI_VAL)
    return d
def _viet_dataframe(data,*args,**kwargs): return _orig_dataframe(_vi_df(data),*args,**kwargs)
st.dataframe=_viet_dataframe

DEFAULT=Path(__file__).with_name('Real_Estate_Project_Master.xlsx')
upload=st.sidebar.file_uploader('Tải file Master Excel', type=['xlsx'])
src=upload if upload else DEFAULT

@st.cache_data(show_spinner=False)
def load_book(source):
    x=pd.ExcelFile(source)
    return {s:pd.read_excel(source,sheet_name=s,header=2) for s in x.sheet_names}

try:
    book=load_book(src)
except Exception as e:
    st.error(f'Không đọc được file Master: {e}')
    st.stop()

def tab(name):
    if name not in book: return pd.DataFrame()
    return book[name].dropna(how='all').copy()

portfolio=tab('01_Danh_muc_du_an')
assump=tab('17_Gia_dinh_mo_hinh')
health=tab('16_Suc_khoe_du_an')
project_codes=portfolio['Mã dự án'].dropna().astype(str).tolist()

# --------------------- MODEL ENGINE ---------------------
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
    b1,b2,b3,b4=st.tabs(['B01-DN • CĐKT','B02-DN • KQKD','B03-DN • LCTT','B04-DN • Thuyết minh'])
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

# Sidebar navigation
mode=st.sidebar.radio('Cấp điều hành',['Toàn công ty / Danh mục','Chi tiết một dự án'])
selected=None
if mode=='Chi tiết một dự án':
    label={r['Mã dự án']:f"{r['Mã dự án']} – {r['Tên dự án']}" for _,r in portfolio.iterrows()}
    selected=st.sidebar.selectbox('Chọn dự án',project_codes,format_func=lambda x:label.get(x,x))
st.sidebar.caption('V10 kế thừa lõi V9 và bổ sung lớp dữ liệu vận hành thực tế: nhập/ánh xạ dữ liệu, khu vực chờ, đối chiếu, nhật ký thay đổi, phân quyền, khóa kỳ và kiểm soát sẵn sàng vận hành.')

# Scenario controls used in both levels
with st.sidebar.expander('Kịch bản điều hành V10', expanded=False):
    price_change=st.slider('Giá bán',-30,20,0,1)/100
    absorption_change=st.slider('Tốc độ hấp thụ',-50,30,0,5)/100
    cost_change=st.slider('Chi phí đầu tư/xây dựng',-10,30,0,1)/100
    rate_change=st.slider('Lãi suất vay',-2.0,5.0,0.0,0.25)/100
    delay_add=st.slider('Chậm pháp lý bổ sung (ngày)',0,365,0,15)
    credit_limit_change=st.slider('Thay đổi hạn mức tín dụng',-50,30,0,5)/100
scen={'price_change':price_change,'absorption_change':absorption_change,'cost_change':cost_change,'rate_change':rate_change,'delay_days_add':delay_add,'credit_limit_change':credit_limit_change}

model_all=all_models(scen)

# --------------------- PORTFOLIO ---------------------
if mode=='Toàn công ty / Danh mục':
    entities=tab('29_Danh_muc_phap_nhan')
    treasury=tab('32_Treasury_Cash_Pool')
    debtbook=tab('33_No_vay_Covenant')
    elimin=tab('31_Loai_tru_hop_nhat')
    waterfall=tab('34_Waterfall_Nguon_von')
    consdata=tab('35_BCTC_Hop_nhat_Data')
    ceo_actions=tab('39_Quyet_dinh_TGD_V7')

    tabs=st.tabs([
        '01. Bàn điều hành tập đoàn','02. Pháp nhân & SPV','03. Ngân quỹ / Điều phối tiền',
        '04. Nợ vay & Cam kết tài chính','05. BCTC hợp nhất','06. Loại trừ nội bộ',
        '07. Thác nguồn vốn','08. Dòng tiền 60T & Kiểm tra sức chịu đựng','09. Sức khỏe & EWS',
        '10. Quyết định TGĐ','11. Hiệu suất & Dự báo','12. Phân bổ vốn',
        '13. Bộ báo cáo HĐQT/CFO','14. Kế hoạch & Phiên bản','15. Khóa sổ tháng',
        '16. Ước tính mới nhất & Chênh lệch','17. Vốn lưu động & Thuế','18. Quản trị dữ liệu',
        '19. Nhập dữ liệu & Ánh xạ','20. Đối chiếu dữ liệu','21. Nhật ký & Phân quyền',
        '22. Khóa kỳ & Luồng dữ liệu','23. Kiểm soát vận hành','24. Pháp lý & chính sách'
    ])

    with tabs[0]:
        total_inv=pd.to_numeric(portfolio['Tổng mức đầu tư (tỷ)'],errors='coerce').sum()
        total_rev=pd.to_numeric(portfolio['Doanh thu kỳ vọng (tỷ)'],errors='coerce').sum()
        group_gap=pd.to_numeric(treasury.loc[treasury['Mã pháp nhân'].astype(str)!='SPV06','Khoảng thiếu vốn 12T (tỷ)'],errors='coerce').sum() if len(treasury) else 0
        bank_debt=pd.to_numeric(debtbook['Dư nợ hiện tại (tỷ)'],errors='coerce').sum() if len(debtbook) else 0
        cov_warn=(debtbook['Trạng thái'].astype(str)=='Cảnh báo').sum() if len(debtbook) else 0
        c1,c2,c3,c4,c5=st.columns(5)
        c1.metric('Pháp nhân hợp nhất',5)
        c2.metric('Tổng mức đầu tư',f'{total_inv:,.0f} tỷ')
        c3.metric('Khoảng thiếu vốn 12T',f'{group_gap:,.0f} tỷ')
        c4.metric('Dư nợ ngân hàng',f'{bank_debt:,.0f} tỷ')
        c5.metric('Cam kết tài chính cảnh báo',int(cov_warn))

        dash=portfolio.merge(health[['Mã dự án','Điểm sức khỏe','Xếp loại']],on='Mã dự án',how='left',suffixes=('','_EWS'))
        if len(treasury):
            tgap=treasury[['Mã dự án','Khoảng thiếu vốn 12T (tỷ)','Cảnh báo']].copy()
            tgap=tgap[tgap['Mã dự án'].astype(str)!='-']
            dash=dash.merge(tgap,on='Mã dự án',how='left')
        st.dataframe(dash[['Mã dự án','Tên dự án','Giai đoạn','Điểm sức khỏe_EWS','Mức cảnh báo','Khoảng thiếu vốn 12T (tỷ)','Cảnh báo','Vấn đề chính']],use_container_width=True,hide_index=True)

        if len(consdata):
            plot=consdata[['Năm mô hình','Doanh thu hợp nhất','LNST hợp nhất']].copy()
            st.plotly_chart(px.bar(plot,x='Năm mô hình',y=['Doanh thu hợp nhất','LNST hợp nhất'],barmode='group',title='Doanh thu & LNST hợp nhất sau loại trừ'),use_container_width=True)
        st.info('V9: SPV06/DA06 đang M&A tiếp tục được để ngoài BCTC hợp nhất và Điều phối tiền cho đến khi hoàn tất giao dịch.')

    with tabs[1]:
        st.subheader('Cấu trúc công ty mẹ – SPV – dự án')
        st.dataframe(entities,use_container_width=True,hide_index=True)
        st.caption('Tỷ lệ sở hữu và phương pháp hợp nhất là dữ liệu đầu vào quản trị; khi áp dụng thật cần khớp hồ sơ pháp lý và BCTC từng pháp nhân.')

    with tabs[2]:
        st.subheader('Ngân quỹ & Điều phối tiền')
        st.dataframe(treasury,use_container_width=True,hide_index=True)
        if len(treasury):
            eligible=treasury[treasury['Được Điều phối tiền?'].astype(str)=='Có']
            surplus=pd.to_numeric(eligible['Khả năng chuyển về CTM'],errors='coerce').sum()
            deficit=pd.to_numeric(eligible['Nhu cầu nhận từ CTM'],errors='coerce').sum()
            c1,c2,c3=st.columns(3)
            c1.metric('Thặng dư có thể điều phối',f'{surplus:,.0f} tỷ')
            c2.metric('Nhu cầu nhận vốn',f'{deficit:,.0f} tỷ')
            c3.metric('Thiếu hụt sau pool',f'{max(0,deficit-surplus):,.0f} tỷ')
        st.warning('Điều phối tiền chỉ là đề xuất điều hành. Không tự động chuyển tiền nếu facility/covenant, tài khoản kiểm soát hoặc mục đích sử dụng vốn hạn chế.')

    with tabs[3]:
        st.subheader('Danh mục nợ và covenant')
        st.dataframe(debtbook,use_container_width=True,hide_index=True)
        if len(debtbook):
            warns=debtbook[debtbook['Trạng thái'].astype(str)=='Cảnh báo']
            if len(warns):
                st.error('Khoản vay cần xử lý: '+', '.join(warns['Mã khoản vay'].astype(str)))
            fig=px.scatter(debtbook,x='DSCR dự báo',y='LTV hiện tại',size='Dư nợ hiện tại (tỷ)',color='Trạng thái',hover_name='Mã khoản vay',title='Bản đồ cam kết tài chính: DSCR – LTV')
            st.plotly_chart(fig,use_container_width=True)

    with tabs[4]:
        st.subheader('BCTC hợp nhất kế hoạch sau loại trừ nội bộ')
        b1=tab('36_B01_DN_Hop_nhat'); b2=tab('37_B02_DN_Hop_nhat'); b3=tab('38_B03_DN_Hop_nhat')
        f1,f2,f3=st.tabs(['B01-DN • CĐKT hợp nhất','B02-DN • KQKD hợp nhất','B03-DN • LCTT hợp nhất'])
        with f1: st.dataframe(b1,use_container_width=True,hide_index=True)
        with f2: st.dataframe(b2,use_container_width=True,hide_index=True)
        with f3: st.dataframe(b3,use_container_width=True,hide_index=True)
        if len(consdata):
            st.caption('Đối chiếu dữ liệu hợp nhất:')
            st.dataframe(consdata,use_container_width=True,hide_index=True)
        st.info('V7 loại SPV06 khỏi hợp nhất; loại trừ doanh thu/chi phí, phải thu/phải trả và cho vay/vay nội bộ. Goodwill và bút toán hợp nhất thực tế cần nhập khi có dữ liệu.')

    with tabs[5]:
        st.subheader('Giao dịch nội bộ & bút toán loại trừ')
        ic=tab('30_Giao_dich_noi_bo')
        st.dataframe(ic,use_container_width=True,hide_index=True)
        st.dataframe(elimin,use_container_width=True,hide_index=True)
        if len(elimin) and (elimin['Trạng thái'].astype(str)!='OK').any():
            st.error('Có chênh lệch giao dịch nội bộ chưa khớp; cần đối chiếu trước khi dùng BCTC hợp nhất.')

    with tabs[6]:
        st.subheader('Thác nguồn vốn')
        psel=st.selectbox('Chọn dự án để xem waterfall',project_codes,key='wf_project')
        wf=waterfall[waterfall['Mã dự án'].astype(str)==psel].copy() if len(waterfall) else waterfall
        st.dataframe(wf,use_container_width=True,hide_index=True)
        if len(wf):
            shortage=pd.to_numeric(wf['Thiếu hụt sau waterfall'],errors='coerce').min()
            if shortage>0: st.error(f'Còn thiếu {shortage:,.0f} tỷ sau khi sử dụng các nguồn trong waterfall.')
            else: st.success('Waterfall hiện có thể bao phủ nhu cầu theo dữ liệu đầu vào.')

    with tabs[7]:
        agg=model_all.groupby('Tháng',as_index=False)[['Thu khách hàng (tỷ)','Chi đầu tư (tỷ)','Vay tăng (tỷ)','Trả nợ (tỷ)','Dư nợ cuối kỳ (tỷ)','Funding gap (tỷ)','Tiền cuối kỳ (tỷ)']].sum()
        fig=go.Figure()
        fig.add_trace(go.Scatter(x=agg['Tháng'],y=agg['Thu khách hàng (tỷ)'],name='Thu khách hàng'))
        fig.add_trace(go.Scatter(x=agg['Tháng'],y=agg['Chi đầu tư (tỷ)'],name='Chi đầu tư'))
        fig.add_trace(go.Scatter(x=agg['Tháng'],y=agg['Dư nợ cuối kỳ (tỷ)'],name='Dư nợ'))
        fig.update_layout(title='Dòng tiền và dư nợ 60 tháng',hovermode='x unified')
        st.plotly_chart(fig,use_container_width=True)
        base=all_models({})
        base_gap=base.groupby('Tháng')['Funding gap (tỷ)'].sum().max()
        stress_gap=model_all.groupby('Tháng')['Funding gap (tỷ)'].sum().max()
        base_ni=base['LNST (tỷ)'].sum(); stress_ni=model_all['LNST (tỷ)'].sum()
        c1,c2=st.columns(2)
        c1.metric('Khoảng thiếu vốn đỉnh',f'{stress_gap:,.0f} tỷ',f'{stress_gap-base_gap:+,.0f}')
        c2.metric('LNST 60T',f'{stress_ni:,.0f} tỷ',f'{stress_ni-base_ni:+,.0f}')

    with tabs[8]:
        h=health.sort_values('Điểm sức khỏe').copy()
        st.dataframe(h,use_container_width=True,hide_index=True)
        st.plotly_chart(px.bar(h,x='Điểm sức khỏe',y='Tên dự án',orientation='h',color='Xếp loại',title='Chỉ số sức khỏe dự án'),use_container_width=True)

    with tabs[9]:
        st.subheader('Agenda quyết định TGĐ')
        st.dataframe(ceo_actions,use_container_width=True,hide_index=True)
        high=ceo_actions[ceo_actions['Mức độ'].astype(str)=='Rất cao'] if len(ceo_actions) else ceo_actions
        if len(high): st.error(f'Có {len(high)} quyết định mức Rất cao cần xử lý.')

    with tabs[10]:
        st.subheader('Quản trị hiệu suất & Dự báo cuốn chiếu')
        bva=tab('42_Budget_vs_Actual'); rf=tab('43_Rolling_Forecast'); cv=tab('44_Covenant_Forecast')
        p1,p2,p3=st.tabs(['Ngân sách so với Thực tế','Dự báo cuốn chiếu 24T','Dự báo cam kết tài chính'])
        with p1:
            st.dataframe(bva,use_container_width=True,hide_index=True)
            if len(bva): st.plotly_chart(px.bar(bva,x='Mã dự án',y=['Doanh thu Budget YTD','Doanh thu Actual YTD'],barmode='group',title='Doanh thu Ngân sách so với Thực tế YTD'),use_container_width=True)
        with p2:
            st.dataframe(rf,use_container_width=True,hide_index=True)
            if len(rf) and (rf['Trigger reforecast'].astype(str)=='REFORECAST').any(): st.warning('Có dự án chạm ngưỡng reforecast – cần khóa giả định mới và trình phê duyệt.')
        with p3:
            st.dataframe(cv,use_container_width=True,hide_index=True)
            if len(cv) and (cv['Xác suất breach'].astype(str)=='Cao').any(): st.error('Có covenant có xác suất breach cao trong horizon dự báo.')

    with tabs[11]:
        st.subheader('Phân bổ vốn – Risk-adjusted Return')
        ca=tab('45_Phan_bo_von')
        st.dataframe(ca,use_container_width=True,hide_index=True)
        if len(ca): st.plotly_chart(px.scatter(ca,x='Risk score',y='RAR score',size='Vốn yêu cầu 12T',color='Quyết định',hover_name='Dự án',title='Ma trận phân bổ vốn'),use_container_width=True)

    with tabs[12]:
        st.subheader('Bộ báo cáo HĐQT/CFO – 3–5 phút')
        pack=tab('55_Management_Report')
        if pack.empty: pack=tab('46_Board_CFO_Pack')
        st.dataframe(pack,use_container_width=True,hide_index=True)
        st.caption('V9 dùng Ước tính mới nhất đang hiệu lực làm số điều hành; Ngân sách là mốc so sánh và Số thực tế đã khóa không bị ghi đè.')

    with tabs[13]:
        st.subheader('Kế hoạch, quản lý phiên bản & luồng phê duyệt')
        vc=tab('47_Quan_ly_Phien_ban'); wf=tab('48_Workflow_Phe_duyet')
        p1,p2=st.tabs(['Quản lý phiên bản','Workflow phê duyệt'])
        with p1:
            st.dataframe(vc,use_container_width=True,hide_index=True)
            if len(vc):
                active=vc[vc['Trạng thái'].astype(str).isin(['Đang sử dụng','Đã khóa'])]
                st.success(f'Có {len(active)} phiên bản đang sử dụng/đã khóa.')
        with p2:
            st.dataframe(wf,use_container_width=True,hide_index=True)
            if len(wf):
                critical=wf[wf['Mức độ'].astype(str)=='Rất cao']
                if len(critical): st.error(f'Có {len(critical)} workflow mức Rất cao cần phê duyệt.')

    with tabs[14]:
        st.subheader('Khóa sổ tháng Control')
        mc=tab('49_Monthly_Close')
        st.dataframe(mc,use_container_width=True,hide_index=True)
        if len(mc):
            c1,c2,c3=st.columns(3)
            c1.metric('Công việc',len(mc))
            c2.metric('Đã hoàn thành',(mc['Trạng thái'].astype(str)=='Hoàn thành').sum())
            c3.metric('Quá hạn',(mc['Trạng thái'].astype(str)=='Quá hạn').sum())
            by=mc.groupby('Mã pháp nhân')['Số ngày trễ'].sum().reset_index()
            st.plotly_chart(px.bar(by,x='Mã pháp nhân',y='Số ngày trễ',title='Ngày trễ khóa sổ theo pháp nhân'),use_container_width=True)

    with tabs[15]:
        st.subheader('Ước tính mới nhất & Chênh lệch Bridge')
        le=tab('50_Latest_Estimate'); vb=tab('53_Variance_Bridge')
        l1,l2=st.tabs(['Ước tính mới nhất năm 2026','Cầu nối chênh lệch'])
        with l1:
            st.dataframe(le,use_container_width=True,hide_index=True)
            if len(le): st.plotly_chart(px.bar(le,x='Mã dự án',y=['Budget DT FY','LE Doanh thu FY'],barmode='group',title='Ngân sách so với Ước tính mới nhất – Doanh thu'),use_container_width=True)
        with l2:
            st.dataframe(vb,use_container_width=True,hide_index=True)
            if len(vb): st.plotly_chart(px.bar(vb,x='Mã dự án',y='Net profit bridge proxy',color='Driver lớn nhất',title='Cầu nối chênh lệch theo yếu tố tác động'),use_container_width=True)

    with tabs[16]:
        st.subheader('Vốn lưu động & Thuế / Obligations Planning')
        wc=tab('51_Working_Capital'); tx=tab('52_Ke_hoach_Thue')
        w1,w2=st.tabs(['Vốn lưu động','Thuế & nghĩa vụ'])
        with w1:
            st.dataframe(wc,use_container_width=True,hide_index=True)
            if len(wc): st.plotly_chart(px.bar(wc,x='Mã dự án',y='Net Vốn lưu động',title='Net Vốn lưu động proxy'),use_container_width=True)
        with w2:
            st.dataframe(tx,use_container_width=True,hide_index=True)
            st.warning('Thuế và nghĩa vụ trong module này là kế hoạch dòng tiền quản trị; phải cập nhật theo giao dịch, hồ sơ dự án và nghĩa vụ pháp lý thực tế.')

    with tabs[17]:
        st.subheader('Quản trị dữ liệu – Single Source of Truth')
        dq=tab('54_Quan_tri_Du_lieu')
        st.dataframe(dq,use_container_width=True,hide_index=True)
        if len(dq):
            bad=dq[dq['Trạng thái DQ'].astype(str)=='CẢNH BÁO']
            if len(bad): st.warning('Data domain cần làm sạch/khóa: '+', '.join(bad['Data domain'].astype(str)))

    with tabs[18]:
        st.subheader('Nhập dữ liệu & Ánh xạ')
        cfg=tab('56_Cau_hinh_Import'); mp=tab('57_Tu_dien_Ánh xạ'); staging=tab('58_Staging_Import')
        d1,d2,d3=st.tabs(['Nguồn dữ liệu','Ánh xạ','Khu vực chờ / Lô dữ liệu'])
        with d1: st.dataframe(cfg,use_container_width=True,hide_index=True)
        with d2: st.dataframe(mp,use_container_width=True,hide_index=True)
        with d3:
            st.dataframe(staging,use_container_width=True,hide_index=True)
            if len(staging):
                blocked=(staging['Trạng thái batch'].astype(str)=='Bị chặn').sum()
                if blocked: st.error(f'Có {blocked} lô dữ liệu bị chặn, chưa được chuyển vào Master.')

    with tabs[19]:
        st.subheader('Đối chiếu dữ liệu – nguồn ↔ Master ↔ BCTC')
        rec=tab('59_Doi_chieu_Du_lieu')
        st.dataframe(rec,use_container_width=True,hide_index=True)
        if len(rec):
            bad=rec[rec['Trạng thái'].astype(str)=='LỆCH']
            if len(bad): st.error(f'Có {len(bad)} đối chiếu bị lệch cần xử lý trước khi khóa báo cáo.')

    with tabs[20]:
        st.subheader('Nhật ký thay đổi & Kiểm soát truy cập')
        aud=tab('60_Audit_Trail'); roles=tab('61_Phan_quyen')
        a1,a2=st.tabs(['Nhật ký thay đổi','Phân quyền / SoD'])
        with a1:
            st.dataframe(aud,use_container_width=True,hide_index=True)
            if len(aud):
                pending=(aud['Trạng thái review'].astype(str)=='Chờ review').sum()
                if pending: st.warning(f'{pending} thay đổi đang chờ review.')
        with a2: st.dataframe(roles,use_container_width=True,hide_index=True)

    with tabs[21]:
        st.subheader('Khóa kỳ & Luồng dữ liệu')
        lock=tab('62_Khoa_Ky'); pipe=tab('63_Data_Luồng dữ liệu')
        p1,p2=st.tabs(['Khóa kỳ','Luồng dữ liệu'])
        with p1: st.dataframe(lock,use_container_width=True,hide_index=True)
        with p2:
            st.dataframe(pipe,use_container_width=True,hide_index=True)
            if len(pipe):
                late=(pipe['SLA status'].astype(str)=='Quá SLA').sum()
                if late: st.error(f'Có {late} luồng dữ liệu quá SLA.')

    with tabs[22]:
        st.subheader('Sẵn sàng vận hành / Kiểm thử / Kiểm soát vận hành')
        ready=tab('64_GoLive_Readiness'); ops=tab('65_Operational_Control')
        g1,g2=st.tabs(['Mức sẵn sàng vận hành','Kiểm soát vận hành'])
        with g1:
            st.dataframe(ready,use_container_width=True,hide_index=True)
        with g2:
            st.dataframe(ops,use_container_width=True,hide_index=True)
            if len(ops):
                row=ops[ops['KPI kiểm soát'].astype(str)=='Go-Live Decision']
                if len(row) and str(row.iloc[0]['Giá trị'])=='NO-GO':
                    st.error('Trạng thái hiện tại: CHƯA VẬN HÀNH. Cần đóng các điều kiện trọng yếu trước khi vận hành vận hành thực tế.')
                elif len(row):
                    st.success('Trạng thái hiện tại: SẴN SÀNG VẬN HÀNH.')

    with tabs[23]:
        laws=tab('12_Cap_nhat_phap_ly')
        st.dataframe(laws,use_container_width=True,hide_index=True,column_config={'Nguồn chính thức':st.column_config.LinkColumn('Nguồn chính thức')})
        st.warning('Các dòng “Sắp có hiệu lực/Policy watch” không được dùng như nghĩa vụ hiện hành. Hồ sơ thực tế phải rà văn bản gốc, quy định địa phương và tình trạng cụ thể của dự án.')

# --------------------- PROJECT ---------------------
else:
    pr=portfolio[portfolio['Mã dự án'].astype(str)==selected].iloc[0]
    a=assump[assump['Mã dự án'].astype(str)==selected].iloc[0]
    pm=build_project_model(a,scen); pa=annualize(pm)
    st.subheader(f"{pr['Mã dự án']} – {pr['Tên dự án']}")
    st.caption(f"{pr['Địa phương']} | {pr['Giai đoạn']} | {pr['Loại hình']}")
    tabs=st.tabs(['01. Tổng quan','02. Mô hình 60T','03. BCTC Việt Nam','04. Truyền dẫn pháp lý','05. Pháp lý','06. Đường găng','07. Phòng hồ sơ','08. Chi đầu tư (CAPEX)','09. Bán hàng','10. Nguồn vốn','11. Khả thi/M&A','12. Rủi ro & EWS','13. Việc cần làm','14. Checklist pháp lý','15. Thực tế & Ước tính mới nhất','16. Vốn lưu động & Thuế'])
    with tabs[0]:
        hs=health[health['Mã dự án'].astype(str)==selected]; score=safe_num(hs['Điểm sức khỏe'].iloc[0]) if len(hs) else 0
        gap=pm['Funding gap (tỷ)'].max(); peak=pm['Dư nợ cuối kỳ (tỷ)'].max(); ni=pm['LNST (tỷ)'].sum()
        c1,c2,c3,c4,c5=st.columns(5); c1.metric('Sức khỏe',f'{score:.0f}/100'); c2.metric('Pháp lý',f"{safe_num(pr['Tiến độ pháp lý']):.0%}"); c3.metric('Dư nợ đỉnh',f'{peak:,.0f} tỷ'); c4.metric('Khoảng thiếu vốn đỉnh',f'{gap:,.0f} tỷ'); c5.metric('LNST 60T',f'{ni:,.0f} tỷ')
        st.warning(f"Vấn đề chính: {pr['Vấn đề chính']}")
        if gap>0: st.error('LEGAL/FINANCE GATE: kịch bản hiện tại tạo funding gap; cần xử lý tiến độ, vốn hoặc chính sách bán hàng trước khi phê duyệt kế hoạch.')
    with tabs[1]:
        fig=go.Figure(); fig.add_trace(go.Scatter(x=pm['Tháng'],y=pm['Thu khách hàng (tỷ)'],name='Thu khách hàng')); fig.add_trace(go.Scatter(x=pm['Tháng'],y=pm['Chi đầu tư (tỷ)'],name='Chi đầu tư')); fig.add_trace(go.Scatter(x=pm['Tháng'],y=pm['Dư nợ cuối kỳ (tỷ)'],name='Dư nợ')); fig.update_layout(title='Dòng tiền – nợ vay 60 tháng',hovermode='x unified'); st.plotly_chart(fig,use_container_width=True)
        st.dataframe(pm,use_container_width=True,hide_index=True)
    with tabs[2]:
        st.dataframe(pa,use_container_width=True,hide_index=True)
        st.plotly_chart(px.bar(pa,x='Năm mô hình',y=['Doanh thu (tỷ)','LNST (tỷ)'],barmode='group',title='KQKD quản trị 5 năm'),use_container_width=True)
        st.caption('Bảng cân đối quản trị được theo dõi qua tiền, dư nợ và CIP/hàng tồn kho proxy trong mô hình; không thay thế BCTC kế toán.')
    with tabs[3]:
        delay=safe_num(a['Chậm pháp lý cơ sở (ngày)'])+delay_add; shift=int(np.ceil(delay/30)); extra_interest=safe_num(a['Dư nợ đầu kỳ (tỷ)'])*(safe_num(a['Lãi suất vay/năm'])+rate_change)/12*shift
        c1,c2,c3,c4=st.columns(4); c1.metric('Chậm pháp lý',f'{delay:.0f} ngày'); c2.metric('Dịch mở bán',f'{shift} tháng'); c3.metric('Lãi vay tăng proxy',f'{extra_interest:,.0f} tỷ'); c4.metric('Khoảng thiếu vốn đỉnh',f"{pm['Funding gap (tỷ)'].max():,.0f} tỷ")
        st.markdown('**Chuỗi truyền dẫn:** Pháp lý → mở bán/ghi nhận → thu tiền → nhu cầu vay → lãi vay → LNST/funding gap.')
    with tabs[4]: st.dataframe(tab('02_Phap_ly').query('`Mã dự án` == @selected'),use_container_width=True,hide_index=True)
    with tabs[5]:
        cp=tab('14_Duong_gang_PL'); cp=cp[cp['Mã dự án'].astype(str)==selected].copy();
        for c in ['Ngày bắt đầu KH','Ngày kết thúc dự báo']: cp[c]=pd.to_datetime(cp[c],errors='coerce')
        if len(cp):
            fig=px.timeline(cp,x_start='Ngày bắt đầu KH',x_end='Ngày kết thúc dự báo',y='Mốc/đầu việc',color='Trạng thái',hover_data=['Mốc khóa','Ảnh hưởng khi chậm']); fig.update_yaxes(autorange='reversed'); st.plotly_chart(fig,use_container_width=True)
        st.dataframe(cp,use_container_width=True,hide_index=True); st.error('LEGAL GATE: chỉ chuyển mốc khi đủ toàn bộ điều kiện pháp lý áp dụng, không dựa riêng vào % tiến độ.')
    with tabs[6]: st.dataframe(tab('15_Phong_ho_so_PL').query('`Mã dự án` == @selected'),use_container_width=True,hide_index=True)
    with tabs[7]:
        d=tab('03_CAPEX'); d=d[d['Mã dự án'].astype(str)==selected]; st.dataframe(d,use_container_width=True,hide_index=True)
        if len(d): st.plotly_chart(px.bar(d,x='Hạng mục',y=['Ngân sách (tỷ)','Dự báo hoàn thành'],barmode='group',title='CAPEX: ngân sách và EAC'),use_container_width=True)
    with tabs[8]: st.dataframe(tab('04_Ban_hang').query('`Mã dự án` == @selected'),use_container_width=True,hide_index=True)
    with tabs[9]: st.dataframe(tab('05_Nguon_von').query('`Mã dự án` == @selected'),use_container_width=True,hide_index=True)
    with tabs[10]: st.dataframe(tab('08_Feasibility_MA').query('`Mã dự án` == @selected'),use_container_width=True,hide_index=True)
    with tabs[11]:
        st.dataframe(tab('06_Rui_ro').query('`Mã dự án` == @selected'),use_container_width=True,hide_index=True)
        if len(hs): st.dataframe(hs,use_container_width=True,hide_index=True)
    with tabs[12]: st.dataframe(tab('07_Hanh_dong').query('`Mã dự án` == @selected'),use_container_width=True,hide_index=True)
    with tabs[13]: st.dataframe(tab('13_Checklist_phap_ly').query('`Mã dự án` == @selected'),use_container_width=True,hide_index=True)
    with tabs[14]:
        act=tab('41_Actual_Theo_thang'); act=act[act['Mã dự án'].astype(str)==selected] if len(act) else act
        le=tab('50_Latest_Estimate'); le=le[le['Mã dự án'].astype(str)==selected] if len(le) else le
        st.dataframe(le,use_container_width=True,hide_index=True)
        if len(act):
            st.plotly_chart(px.line(act,x='Tháng',y=['Doanh thu Actual (tỷ)','Thu tiền Actual (tỷ)','CAPEX Actual (tỷ)'],markers=True,title='Số thực tế theo tháng'),use_container_width=True)
            st.dataframe(act,use_container_width=True,hide_index=True)
    with tabs[15]:
        wc=tab('51_Working_Capital'); wc=wc[wc['Mã dự án'].astype(str)==selected] if len(wc) else wc
        tx=tab('52_Ke_hoach_Thue'); tx=tx[tx['Mã dự án'].astype(str)==selected] if len(tx) else tx
        st.dataframe(wc,use_container_width=True,hide_index=True)
        st.dataframe(tx,use_container_width=True,hide_index=True)
        st.caption('Kế hoạch thuế/nghĩa vụ là module quản trị dòng tiền, không thay thế xác định nghĩa vụ thuế/pháp lý chính thức.')
