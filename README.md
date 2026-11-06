# Office Suite V1 🚀

An open-source, lightweight, and localized productivity suite built from scratch to replace bulky and paid office applications. Combining a modern frontend with a high-performance **FastAPI** backend, this suite delivers an all-in-one seamless workflow for Documents, Spreadsheets, and Slides.

*This project is dedicated to my country, Pakistan. 🇵🇰*

---

## 🔥 Key Features

### 🎛️ All-in-One Dashboard
* **Multi-Tab Workflow:** Switch between **Word, Sheets, and Slides** instantly inside a single application window without heavy system resource overhead.
* **Localized UX:** Designed with custom native placeholders (e.g., *"Yahan likhna shuru karein..."*) making it extremely accessible for local users.

### ✍️ Advanced Rich Text Editor (Word)
* **Custom Typography:** Supports over **10 professional font formats** (including Georgia, Arial, Times New Roman, Calibri, etc.) with real-time size adjustments.
* **Flexible Formatting:** Dedicated controls for Headings, Subheadings, Paragraph styling, Text Color picking, Bold, Italic, and Underline attributes.
* **Dynamic Table Matrix:** Built-in modular system to inject, append, or delete rows and columns dynamically from the frontend toolbar.

### 📊 Optimized Spreadsheet (Sheets)
* **Lightweight Matrix Design:** Features a dedicated data grid optimized specifically for quick data entries and tracking.
* **Fixed Data Horizon (V1 Scope):** Pre-allocated for a **100-row micro-matrix** to ensure instantaneous rendering, low memory usage, and zero lag on entry-level hardware during file operations.

### 📂 Broad File System Integration
* **8+ Export Formats:** Create, open, and save files seamlessly into multiple extensions including `.docx`, `.xlsx`, `.md` (Markdown), `.html`, `.txt`, and more.
* **Flexible Directory Explorer:** Fully integrated with native OS file dialogs—allowing users to **Open** or **Save** files anywhere in their custom filesystem paths.

---

## 🛠️ Tech Stack & Architecture

This suite utilizes a secure **Client-Server Architecture** running completely offline on your local machine:
* **Frontend:** Modular HTML5, CSS3 (Minimalist UI layout), and Vanilla JavaScript for fast rendering.
* **Backend Core:** **FastAPI (Python)** — chosen for high performance, asynchronous request handling, and robust data validation.
* **WSGI/ASGI Server:** Powered by **Uvicorn/Gunicorn** utilizing an industrial-grade **8-worker process management setup** to ensure zero lag during high-compute file conversions.

---

## 🚀 Installation & Setup

### Automated Booting
Simply execute the batch automation file included in the root directory:
```bash
Build_OfficeSuite_EXE.bat
```

### Manual Fallback Installation
If your system blocks automation scripts, the framework provides an intelligent live-error handler on the terminal screen. Follow these steps to spin up the server manually:

1. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
   *(Or manually fetch required modules: `pip install fastapi uvicorn python-docx openpyxl`)*

2. **Launch the Core FastAPI Server:**
   ```bash
   python office_server.py
   ```
3. **Access the Suite:** Open your web browser and navigate to `http://127.0.0.1:8000`.

---

## 📅 What's Next? (Roadmap)
* [ ] **Next-Gen LaTeX Integration:** Shifting the documentation engine toward LaTeX for high-fidelity engineering reports and flawless mathematical formula compilation.
* [ ] **Dynamic Sheet Scalability:** Expanding the spreadsheet layout from a fixed 100-row micro-matrix to infinite virtual scrolling for massive data manipulation.
* [ ] **Local PDF Compilation:** Direct file conversion pipeline from rich-text/markdown assets directly into standard production-quality PDFs.

---

## 📄 License
Distributed under the **GPL-3.0 License**. Feel free to inspect, fork, and scale this project!
# Office Suit version 1
