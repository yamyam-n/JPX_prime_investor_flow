# Verification notes for v1.8

- JPX official PDF URL confirmed for 2026/1 week4 (Prime, value):
  `.../vk0khi000000qsw6-att/stock_val_1_260104.pdf`
- The same path with `.xls` was confirmed by the web fetch layer to return `application/vnd.ms-excel`.
- JPX official PDF values for Prime / Foreigners / 2026/1 week4 current-week column:
  - Sales: 23,762,871,033 thousand yen
  - Purchases: 23,949,897,312 thousand yen
  - Balance: 187,026,279 thousand yen = 1,870.26279 oku yen
- v1.8 parser self-test reproduces the balance from the published JPX layout and passes `Balance = Purchases - Sales` validation.

The runtime app additionally GETs each Excel candidate, verifies Excel magic bytes/content type, opens it with xlrd/openpyxl, selects the Prime sheet, and rejects data that fails consistency checks.
