# Office Suite

Ek hi application mein **Word**, **Sheet (Excel)** aur **Slides (PowerPoint)**.
Python server background mein chalta hai aur app ki apni window (GUI) saath khulti hai.

## Files

| File | Kaam |
|---|---|
| `office_server.py` | Python backend + launcher (server aur window dono start karta hai) |
| `OfficeSuite.html` | App ka poora interface (Word, Sheet, Slides) |
| `Start_OfficeSuite.bat` | Windows par ek click se chalane ke liye |
| `Build_EXE.bat` | Asli `.exe` application banane ke liye |

Teeno/chaaron files **ek hi folder** mein rakhein.

## Features

- **Word:** headings (Heading, Sub Heading, Small Heading), font change, font size 14 se 30, bold, italic, underline, rang, alignment, bullets, numbering, **tables** (+Row, -Row, +Col, -Col, Del Table), undo/redo.
- **Sheet:** Excel jaisa grid, formulas: `=A1+B1`, `=SUM(A1:A5)`, `AVERAGE`, `MIN`, `MAX`, `COUNT`.
- **Slides:** slides add, duplicate, delete, upar/neeche, 5 themes, full screen **Present** mode (arrow keys, Esc).
- **New / Open / Save:** nayi file banana, purani file kholkar edit karna, aur save.

## Supported formats

| Tab | Open | Save |
|---|---|---|
| Word | `.docx` `.txt` `.html` `.md` | `.docx` `.txt` `.html` |
| Sheet | `.xlsx` `.csv` `.tsv` `.txt` | `.xlsx` `.csv` `.txt` |
| Slides | `.pptx` `.txt` | `.pptx` `.txt` |

Purani `.doc`, `.xls`, `.ppt` files support nahi. Unhe Office mein `.docx`, `.xlsx`, `.pptx` save karke kholein.

## Chalane ka tareeka

**Requirement:** Python 3.9 ya naya.

1. Saari files ek folder mein rakhein.
2. Windows par `Start_OfficeSuite.bat` par double-click karein.
   Ya terminal mein: `python office_server.py`
3. Pehli baar libraries khud install hoti hain (`fastapi`, `uvicorn`, `pywebview`, `python-docx`, `openpyxl`, `python-pptx`). Is waqt internet chahiye.
4. Server start hone ke baad app ki window khul jati hai. **Window band karne par server bhi band ho jata hai.**
5. Agar window wali library (`pywebview`) na chale, to app browser mein khulta hai. Tab band karne ke liye terminal mein `Ctrl+C` dabayein.

Manual install (agar auto-install na ho):

```
pip install fastapi "uvicorn[standard]" pywebview python-docx openpyxl python-pptx
```

## Asli .exe application banana (Windows)

1. `Build_EXE.bat` par double-click karein (internet chahiye, kuch minute lagte hain).
2. Taiyar app milega: `dist\OfficeSuite\OfficeSuite.exe`.
3. Doosre computer par dene ke liye poora `dist\OfficeSuite` folder zip karke bhejein (wahan Python ki zaroorat nahi).
4. Installer (Setup.exe) chahiye to us folder ko **Inno Setup** mein use karein.
5. Icon ke liye `Build_EXE.bat` ki pyinstaller command mein `--icon app.ico` add karein.

Notes:
- Exe mein server ek process se chalta hai. Multi-worker tezi sirf `python office_server.py` mode mein milti hai.
- Naye PyInstaller exe par antivirus kabhi galat warning de sakta hai.

## Settings (optional)

Chalane se pehle environment variables set karein, jaise `set WORKERS=4` (Windows) ya `WORKERS=4 python office_server.py` (Mac/Linux).

| Variable | Default | Matlab |
|---|---|---|
| `WORKERS` | CPU ke hisaab se (max 8) | Server processes ki ginti |
| `MAX_ROWS` | `300` | Sheet mein bari file ki kitni rows dikhani hain |
| `MAX_UPLOAD_MB` | `100` | Open ki jane wali file ka max size |
| `PORT` | `8000` | Server ka port (busy ho to agla free port) |
| `HOST` | `127.0.0.1` | Sirf isi computer ke liye. `0.0.0.0` se LAN par khul jata hai |
| `SAVE_COPIES` | `1` | `0` karne par `exports` folder mein copy nahi banti |

## Bade data ke baare mein

- Bari `.xlsx` (5 MB se upar) aur `.csv` files **stream** karke padhi jati hain. Sirf shuru ki `MAX_ROWS` rows aur 40 columns screen par aate hain, aur "bari file" ka message dikhta hai.
- Test: 3 lakh rows wali (13 MB) xlsx kuch millisecond mein khuli.
- Sheet ka grid 100 rows aur 26 columns se shuru hota hai. `MAX_ROWS` bahut badhane par grid bhari ho jata hai.
- Server tez chalne ke liye banaya gaya hai (page memory mein cache, gzip, multi-worker), lekin asli speed aapke computer ki CPU, RAM aur network par depend karti hai. "Zero error" ki guarantee koi software nahi deta.

## Bina server ke

`OfficeSuite.html` ko seedha browser mein kholne par bhi **save** aur **open** chalte hain (internet se file banane wali libraries load hoti hain). Is mode mein:
- `.pptx` file **kholne** ke liye server zaroori hai.
- `.docx` kholne par sirf text, headings, lists aur tables aati hain (rang aur font kam).

## Save kaise hota hai

- Chrome / Edge mein "Save" dabane par aap file ki jagah khud chun sakte hain.
- Baaki browsers mein file seedha Downloads folder mein jati hai.
- Server chal raha ho to har save ki ek copy `exports` folder mein bhi banti hai.

## Masle aur hal

| Masla | Hal |
|---|---|
| `fastapi install nahi hua` | Terminal mein upar wali `pip install` command chalayein, internet check karein |
| Window nahi khulti, browser khulta hai | `pip install pywebview`. Linux par GTK/Qt bhi chahiye ho sakta hai |
| Port busy | App khud agla free port le leta hai, ya `PORT=9000` set karein |
| File khulte waqt error | Check karein ke file kharab ya password-protected to nahi |
| Pehli baar save/open mein dair | Libraries load ho rahi hain, doosri baar tez hoga |

## Jo abhi nahi hai

- Images, charts, table mein cell merge, slides mein images/animations.
- Purani `.doc`, `.xls`, `.ppt` files.
- 10 lakh rows ka poora grid ek saath dikhana (abhi sirf shuru ki rows dikhti hain).
