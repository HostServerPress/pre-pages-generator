# 📄 Pre-Pages Generator & Visualiser

An automated web application built for MMU Press to streamline the creation of academic journal pre-pages (front matter) and visualize publication analytics. The system integrates directly with Open Journal Systems (OJS) CSV exports, Crossref APIs, and cloud databases to eliminate manual data entry.

## ✨ Key Features

### 📑 Document Generator (Tab 1)
* **Dynamic Template Rendering:** Uses Jinja2 tags to inject metadata, editorial boards, and Table of Contents directly into standard Microsoft Word (`.docx`) templates.
* **Smart Cover Image Replacement:** Automatically swaps a placeholder image in the Word template with an uploaded high-resolution cover.
* **Web Scraping Automation:** Fetches and categorizes Editorial Board members and "About the Journal" text directly from live journal websites.
* **OJS CSV Integration:** Parses raw OJS "Articles Report" CSVs, automatically filtering out rejected papers and sorting by scheduled status.
* **Hybrid Page Calculation Engine:**
  * **PDF Extraction:** Calculates starting page numbers by reading the page counts of physical PDF files inside a `.zip` archive.
  * **Crossref API:** Automatically fetches registered page ranges using article DOIs.
* **Advanced Sorting:** Automatically handles Roman numeral pagination (`i, ii, iii`) for front-matter (Editorials, Prefaces) using a custom `00` sorting logic.

### 📊 Data Visualiser (Tab 2)
* **Executive KPIs:** Displays total submissions, conversion rates (active vs. rejected), and unique author counts.
* **Interactive Analytics:** Generates dynamic Pie Charts (Section breakdown) and Bar Charts (Top publishing authors) using Plotly.
* **Topic Trends:** Renders an image-based Word Cloud based on article titles to identify trending research topics.

---

## 🛠️ Tech Stack

* **Frontend & Backend Framework:** [Streamlit](https://streamlit.io/)
* **Database (BaaS):** [Supabase](https://supabase.com/) (PostgreSQL)
* **Document Engine:** `docxtpl` (Jinja2 for Microsoft Word)
* **Data Manipulation:** `pandas`
* **Web Scraping:** `BeautifulSoup4`, `requests`
* **PDF Parsing:** `PyPDF2`
* **Visualizations:** `plotly`, `wordcloud`

---

## 📂 Repository Structure

```text
├── fyp_app.py           # Main Streamlit application and wizard logic
├── utils_cleaning.py    # Utility module for background data extraction and text sanitization
├── utils_viz.py         # Utility module containing the Data Visualiser charts and Word Cloud logic
├── requirements.txt     # Python dependencies
└── README.md            # Project documentationgit add README.md