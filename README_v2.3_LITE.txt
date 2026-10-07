PAYROLL SLIP v2.3 LITE

1. IMPORTANT: Code.gs and index.html now use the SAME Web App URL.
2. Code.gs is based on the original working backend; no PayrollConfig/ExtraData layer.
3. PayrollSync.bas keeps the original workflow and converts Salary column mapping to Header Mapping.
4. Salary headers are searched in rows 1-6. If a header is not found, the old column position is used and a warning is written to PayrollSync.log.
5. The data row still starts at row 7. Sheets Salary and DSCNV are unchanged.
6. Deploy Code.gs as Web App, then put the same URL into index.html. The included URL is the current API URL used by VBA.
7. Replace the VBA module PayrollSync.bas. Do not paste the huge code into chat; import the .bas file.

Recommended test: open the web app -> employee login -> view payslip. Then run VBA for one month and confirm Payroll/SyncStatus.
