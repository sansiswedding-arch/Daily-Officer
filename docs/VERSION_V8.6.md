# V8.6 Release Notes

SANSIS DAILY OFFICER — V8.6 ADMIN FINANCE

V8.6 menambahkan menu Finance pada Administrator App.

ADMIN MENU
- Dashboard
- Employee Management
- Attendance Monitoring
- Finance
- Logout

FINANCE ADMIN
Finance memiliki 4 tab:

1. OVERVIEW
Menampilkan:
- Estimated Payment
- Approved Reimbursement
- Approved Claim
- Estimated Total Payable

Tabel per Employee:
- Employee
- Division
- WFO Days
- Payment Estimate
- Approved Reimbursement
- Approved Claim
- Total Payable

Formula:
Estimated Total Payable =
Payment Estimate + Approved Reimbursement + Approved Claim

2. PAYMENT ELEMENTS
Administrator dapat memilih Employee dan mengatur Payment Element.

Calculation Type:
- Reference
  Hanya angka referensi. Tidak masuk Estimated Payment.
- Fixed Payment
  Dihitung penuh setiap periode.
- Daily WFO Rate
  Nominal x jumlah hari WFO pada periode tanggal 5–4.

Administrator dapat:
- Add Payment Element
- Edit Payment Element
- Active / Inactive Payment Element
- Delete Payment Element

3. REIMBURSEMENT
Administrator dapat melihat semua pengajuan Reimbursement Employee:
- Employee
- Date
- Category
- Description
- Amount
- Proof
- Status

Status dapat diubah:
- Pending
- Approved
- Rejected

Bukti dapat dibuka melalui View.

4. KLAIM
Administrator dapat melihat semua Klaim Achievement Employee:
- Employee
- Date
- Claim Name
- Description
- Amount
- Proof
- Status

Status:
- Pending
- Approved
- Rejected

5. FINANCE PERIOD
Finance menggunakan periode:
tanggal 5 sampai tanggal 4 bulan berikutnya.

Tersedia month selector untuk berpindah periode.

6. DASHBOARD
Finance Overview di Dashboard Administrator sekarang memiliki tombol:
Manage Finance

7. API ADMIN BARU
- getAdminFinanceData(sessionToken, monthKey)
- saveAdminFinanceElement(sessionToken, payload)
- deleteAdminFinanceElement(sessionToken, elementId, monthKey)
- updateAdminReimbursementStatus(sessionToken, reimburseId, status, monthKey)
- updateAdminClaimStatus(sessionToken, claimId, status, monthKey)

Semua API dilindungi requireAdminSession_.

8. PREVIEW
Administrator:
Username: admin.sansis
Password: admin123

Employee:
Username: onky.soerya
Password: preview123

9. DEPLOY
1. Replace Code.gs
2. Replace Index.html
3. Run setupSansisDailyOfficer()
4. Deploy Web App sebagai New Version
5. Login Administrator
6. Buka menu Finance

Seluruh fitur V8.5.1 dan Employee App tetap dipertahankan.