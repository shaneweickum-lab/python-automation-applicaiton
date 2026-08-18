# ExcelAutomate 📊⚡
> **In-Browser Python WebAssembly Data & Visual Excel Automation Pipeline (PWA)**

Transform raw business data (CSV, XLSX, JSON, TSV) into executive-ready, beautifully formatted Excel spreadsheets (`.xlsx`) with KPI cards, summary aggregation tables, and visual charts — running **100% in your browser** via Python WebAssembly (Pyodide), with zero backend server required!

---

## ✨ Features

- 🐍 **In-Browser Python WebAssembly Execution**: Powered by Pyodide, executing genuine `pandas` and `openpyxl` client-side in the browser.
- 📁 **Universal Data Ingestion**: Drag & drop CSV, Excel (`.xlsx`, `.xls`), JSON, and TSV files, or load the built-in sample business dataset.
- 📊 **Dynamic Visual Chart Builder**: Switch between Bar, Column, Line, Pie, and Doughnut charts with live Chart.js preview and native embedded Excel charts.
- 🎨 **Executive Styling & Themes**: Choose from Corporate Navy, Modern Emerald, Tech Indigo, and Charcoal Slate palettes with zebra striping, currency number formatting (`$#,##0.00`), auto-filters, and frozen panes.
- 📱 **Installable PWA**: Works offline, installable as a native desktop (macOS/Windows) or mobile app (iOS/Android).
- 🚀 **1-Click Vercel & GitHub Deployment**: Static architecture with preconfigured `vercel.json`.

---

## 🏃 Local Quickstart

You can run the web application locally with any local HTTP server:

```bash
# Python 3 built-in HTTP server
python3 -m http.server 8000
```
Then open your browser to: [http://localhost:8000](http://localhost:8000)

Or run the CLI automation script directly in terminal:
```bash
# Process default or newest file in data/
python3 rawToReadable.py --auto

# Or specify a custom file:
python3 rawToReadable.py --file data/sample_sales_data.csv
```

---

## 🚀 How to Push to GitHub & Deploy on Vercel

### Step 1: Initialize Git and Push to GitHub

```bash
# Initialize git repository
git init

# Add all files
git add .

# Commit changes
git commit -m "Initial commit: In-browser Python Visual Excel Reporter PWA"

# Add your GitHub remote repository (replace with your repo URL)
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git

# Push to main branch
git branch -M main
git push -u origin main
```

### Step 2: Deploy to Vercel (1-Click)

1. Go to [vercel.com](https://vercel.com) and log in with your GitHub account.
2. Click **"Add New..."** -> **"Project"**.
3. Select your GitHub repository.
4. Leave **Framework Preset** as *Other* (Static) and click **"Deploy"**.
5. Vercel will deploy your PWA instantly with global CDN and SSL!

---

## ✉️ SMTP Email Configuration

To enable automated report emailing:
1. In the Web App, navigate to the **Email & Delivery** tab and enter your recipient (`shane@sowedandrooted.com`) and SMTP App Password. Settings are stored strictly in your browser's `localStorage`.
2. For the CLI script, copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

---

## 📄 License
MIT License. Built with Pyodide, OpenPyXL, Pandas, and Chart.js.
