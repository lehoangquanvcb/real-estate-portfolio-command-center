# Real Estate Portfolio Command Center V10

V10 kế thừa V9 và chuyển platform sang kiến trúc **production-ready data & control**.

## Nâng cấp V10
- Cấu hình nguồn dữ liệu ERP / CRM / PMIS / Bank / Legal.
- Từ điển Mapping nguồn → Master/BCTC/KPI.
- Staging & Batch Control trước khi promote vào Master.
- Reconciliation nguồn ↔ Master ↔ BCTC.
- Audit Trail.
- Role / Permission / Segregation of Duties.
- Period Lock: soft close, hard close, controlled adjustment.
- Data Pipeline Control với DQ/SLA/Approval/Promote.
- Go-Live Readiness, UAT, sign-off và Operational Control Center.

## Nguyên tắc
1. Actual đã hard-close không sửa trực tiếp.
2. Batch chỉ được promote khi DQ + reconciliation + approval + period-lock check đạt.
3. Latest Estimate là số điều hành; Budget là baseline; Forecast draft không ghi đè LE.
4. Mọi thay đổi trọng yếu phải có owner, lý do, workflow và audit trail.
5. BCTC/Board Pack chỉ lấy từ dữ liệu/phiên bản đã được phê duyệt.

Dữ liệu hiện tại là bộ dữ liệu mẫu để UAT. Khi triển khai thật, thay bằng extract ERP/PMIS/Bank/CRM/Legal theo mapping trong Master.
