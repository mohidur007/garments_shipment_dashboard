SETUP (once)
1. Put these in the "libs" folder (open each link, press Ctrl+S, same file name):
   https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
   https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js
2. Keep your Excel file in this folder with the name:  data.xlsx
   (to use another name, change FILE at the top of script.js)

SPECIAL SHORTCUT (once per computer)
Browsers block reading files from a folder unless started with one extra setting.
1. Right-click the desktop > New > Shortcut.
2. Paste this as the location (edit the last path to your index.html):
   "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --allow-file-access-from-files --user-data-dir="%USERPROFILE%\EdgeDashboard" "Z:\MRLALON (AQC MTO)\Dashboard\index.html"
3. Name it "Dashboard" and use it every time to open the dashboard.

UPDATING DATA
Edit and save data.xlsx. The page re-reads it every 30 seconds by itself.

EXTRA LIBRARY FOR "Download Excel" (once)
Open https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
Press Ctrl+S and save it in the "libs" folder as exceljs.min.js

EXTRA LIBRARY FOR "Download PDF" (once)
Open https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js
Press Ctrl+S and save it in the "libs" folder as html2canvas.min.js

COMPANY LOGO
Save your logo in the Dashboard folder (next to index.html) with the name logo.png
(logo.jpg also works). It appears in the dashboard header, the printed report, the PDF and the Excel file.
