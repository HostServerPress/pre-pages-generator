import os
import re
import io
import base64
import tempfile
import zipfile
import datetime
import urllib.parse

import requests
import pandas as pd
import streamlit as st
from PIL import Image
from io import BytesIO
from bs4 import BeautifulSoup, NavigableString
from PyPDF2 import PdfReader
from supabase import create_client, Client
from docxtpl import DocxTemplate, RichText

# --- IMPORT CUSTOM MODULES ---
from utils_cleaning import clean_ojs_csv
import utils_viz 

# --- CONFIGURATION ---
DEFAULT_ADDRESS = """JIWE,
Faculty of Computing and Informatics,
Multimedia University,
Persiaran Multimedia,
63100 Cyberjaya, Malaysia.
Website: https://journals.mmupress.com/index.php/jiwe/index
https://doi.org/10.33093/jiwe"""


# ==========================================
#        SUPABASE INTEGRATION
# ==========================================
@st.cache_resource
def init_supabase() -> Client | None:
    """Initializes the Supabase client using Streamlit secrets."""
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception as e:
        st.warning("⚠️ Supabase credentials not found. Running in local session-only mode.")
        return None

supabase = init_supabase()


# ==========================================
#        PART 1: HELPER FUNCTIONS
# ==========================================

def clean_scraped_text(text):
    """
    Sanitizes scraped HTML text to prevent Microsoft Word justification stretching.
    """
    if not text: 
        return ""
        
    # 1. Destroy Non-Breaking Spaces and Zero-Width characters
    text = text.replace('\xa0', ' ').replace('\u200b', '')
    
    # 2. Normalize Windows/Mac line endings
    text = text.replace('\r\n', '\n')
    
    # 3. Separate the text into actual paragraphs (split by double newlines)
    paragraphs = text.split('\n\n')
    
    cleaned_paragraphs = []
    for p in paragraphs:
        # THE MAGIC FIX: Replace single newlines (Soft Returns) inside a paragraph with a normal space
        p = p.replace('\n', ' ')
        
        # Remove any accidental double/triple spaces
        p = re.sub(r'\s+', ' ', p).strip()
        
        if p:
            cleaned_paragraphs.append(p)
            
    # Rejoin the clean text using Hard Paragraph Breaks (double newlines)
    return '\n\n'.join(cleaned_paragraphs)


def make_rich_text(text_input, doc=None):
    """
    Converts raw text into a docxtpl RichText object.
    Applies MS Word justification hacks and dynamically generates native hyperlinks.
    """
    if not text_input: return ""
    rt = RichText()
    
    # 1. Normalize line endings
    clean_input = str(text_input).replace('\r\n', '\n')
    
    # 2. THE MICROSOFT WORD JUSTIFICATION HACK
    clean_input = clean_input.replace('\n', '\t\n')
    
    # 3. Process the bold tags and URLs simultaneously
    parts = clean_input.split("**")
    for i, part in enumerate(parts):
        is_bold = (i % 2 != 0)
        
        # Split the text by URLs (captures https:// or http:// up to the next space)
        url_parts = re.split(r'(https?://[^\s]+)', part)
        
        for u_part in url_parts:
            if not u_part: continue
            
            # If a URL is detected and we have the doc object, make it a true hyperlink!
            if u_part.startswith('http') and doc is not None:
                # Clean off trailing punctuation in case the user typed a period after the link
                clean_url = u_part.rstrip('.,;:"\'')
                trailing_punct = u_part[len(clean_url):]
                
                # Register the URL with the Word Document XML
                url_id = doc.build_url_id(clean_url)
                
                # Add the clickable link in classic Word blue (#0563C1) and underlined
                rt.add(clean_url, url_id=url_id, color='0563C1', underline=True, bold=is_bold)
                
                # Add back any trailing punctuation as normal text
                if trailing_punct:
                    rt.add(trailing_punct, bold=is_bold)
            else:
                rt.add(u_part, bold=is_bold)
                
    return rt
    
def create_bullet_block(text_input):
    """Formats a block of text into individual cleaned lines for rendering."""
    if not text_input: return ""
    lines = str(text_input).split('\n')
    clean_lines = [line.strip() for line in lines if line.strip()]
    full_block = "\n".join(clean_lines)
    return make_rich_text(full_block) 

def int_to_roman(num):
    """Converts an integer to a lowercase Roman numeral (e.g., for front-matter pagination)."""
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syb = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"]
    roman_num = ''
    i = 0
    while num > 0:
        for _ in range(num // val[i]):
            roman_num += syb[i]
            num -= val[i]
        i += 1
    return roman_num

def extract_shortcode_from_csv(df):
    """Searches the URL or DOI columns of the OJS CSV to securely identify the journal shortcode."""
    url_col = next((c for c in df.columns if 'url' in c.lower()), None)
    doi_col = next((c for c in df.columns if 'doi' in c.lower()), None)
    
    shortcode = None
    
    # 1. Try extracting from URL (e.g., .../index.php/jiwe/...)
    if url_col and not df[url_col].empty:
        sample_url = str(df[url_col].iloc[0])
        match = re.search(r'/index\.php/([^/]+)/', sample_url)
        if match: 
            shortcode = match.group(1).lower()
            
    # 2. If URL fails, try extracting from DOI (e.g., 10.33093/jiwe.2022...)
    if not shortcode and doi_col and not df[doi_col].empty:
        sample_doi = str(df[doi_col].iloc[0])
        match = re.search(r'10\.\d+/([^.]+)\.', sample_doi)
        if match: 
            shortcode = match.group(1).lower()
            
    return shortcode

def analyze_template(template_file):
    """
    Reads the Word template OOXML structure, extracts all custom variables, 
    and sorts them in the exact order they appear in the document layout.
    """
    template_file.seek(0)
    
    # 1. Use docxtpl to GUARANTEE we find 100% of the variables (including loops)
    doc = DocxTemplate(template_file)
    all_vars_unordered = doc.get_undeclared_template_variables()
    
    # 2. Read the raw XML to figure out the ORDER
    template_file.seek(0)
    with zipfile.ZipFile(template_file) as z:
        xml = z.read("word/document.xml").decode("utf-8")
        
    clean_xml = re.sub('<[^<]+>', '', xml) 
    
    # 3. Detect which ones are lists (used in loops)
    list_vars_found = re.findall(r'{%\s*(?:p|tr|tc|)\s*for\s+\w+\s+in\s+(\w+)\s*%}', clean_xml)
    
    system_vars = {'journal_name', 'vol', 'issue', 'month', 'year', 'eissn', 
                   'about_text', 'address_text', 'toc_sections'}
    
    # Filter out system variables to get only the custom editorial roles
    role_vars_unordered = [v for v in all_vars_unordered if v not in system_vars]
    
    # 4. SORT the variables based on where their name first appears in the document text
    def get_appearance_index(var_name):
        idx = clean_xml.find(var_name)
        return idx if idx != -1 else 999999 # Push to the bottom if index not found
        
    role_vars_ordered = sorted(role_vars_unordered, key=get_appearance_index)
    
    template_file.seek(0) 
    return role_vars_ordered, list_vars_found

def parse_editorial_html(soup):
    """
    Dynamic Scraper: Identifies ANY heading based on heuristics (bold, h-tags) 
    and groups the subsequent names under that heading to build the Editorial Board.
    """
    data = {}
    
    # Replace <br> tags with a newline character so we can split multiline paragraphs
    for br in soup.find_all("br"):
        br.replace_with("\n")
        
    content_div = soup.find('div', class_='page-header')
    if content_div:
        main_container = content_div.parent 
        elements = main_container.find_all(['p', 'h2', 'h3', 'h4', 'div', 'li'])
    else:
        elements = soup.find_all(['p', 'h2', 'h3', 'h4', 'div', 'li'])

    current_heading = "Uncategorized"
    data[current_heading] = []
    
    # Master list of known roles to help the heuristic engine
    known_roles = [
        'editor-in-chief', 'managing editor', 'co-editor', 
        'support editor', 'advisory board', 'editorial board', 
        'managing editors', 'co-editors', 'support editors'
    ]

    for el in elements:
        text_block = el.get_text(" ", strip=True)
        lines = text_block.split('\n')
        
        for text in lines:
            # Clean up spacing 
            text = text.strip()
            text = re.sub(r'\s+', ' ', text)
            text = text.replace("( ", "(").replace(" )", ")")
            
            if not text or text == "&nbsp;" or text == "Home": continue
            
            text_lower = text.lower()
            
            # Strict exclusion to prevent accidentally deleting organization names
            if text_lower in ["directory of open access journals", "doaj", "at the directory of open access journals (doaj)", "directory of open access journals (doaj)"]: 
                continue
            
            # Inline Role Detector (Catches "Editor-in-Chief: John Doe")
            is_inline_role = False
            for role in known_roles:
                if text_lower.startswith(role + ":") or text_lower.startswith(role + " :"):
                    parts = text.split(":", 1)
                    current_heading = parts[0].strip()
                    if current_heading not in data: data[current_heading] = []
                    
                    person_name = parts[1].strip()
                    if person_name:
                        data[current_heading].append(person_name)
                        
                    is_inline_role = True
                    break 
            
            # If we successfully parsed an inline role, skip the rest of the heuristic checks
            if is_inline_role:
                continue

            # Standard Heuristics for standalone headers
            is_h_tag = el.name in ['h2', 'h3', 'h4']
            is_bold = el.find('strong') is not None or el.find('b') is not None
            is_short = len(text) < 60
            looks_like_person = ',' in text or 'University' in text or 'Universitas' in text or '@' in text or 'Prof' in text or 'Dr' in text or 'Scopus' in text
            
            is_known = text_lower.strip(':') in known_roles
            
            is_likely_header = is_known or ((is_h_tag or (is_bold and is_short)) and not looks_like_person)
            
            if is_likely_header:
                current_heading = text.strip(":")
                if current_heading not in data: data[current_heading] = []
            else:
                clean_item = text.lstrip("•-1234567890. ").strip()
                if clean_item and clean_item != current_heading:
                    data[current_heading].append(clean_item)

    # Clean up and return only categories that actually have names in them
    return {k: v for k, v in data.items() if v}

def parse_about_html(soup):
    """
    Scrapes and cleans the 'About the Journal' section, automatically stopping at specific keywords.
    """
    full_text_lines = []
    content_div = soup.find('div', class_='page-about') or soup.find('div', class_='page_about')
    if not content_div:
        content_div = soup.find('div', class_='pkp_structure_content') or soup.body

    stop_keywords = ["scope limitations", "publication frequency", "open access policy", "peer review process"]

    if content_div:
        for element in content_div.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'li']):
            if element.find_parent(class_=re.compile(r'breadcrumb|nav|menu')): continue
            raw_text = element.get_text(strip=True)
            if raw_text == "Home" or raw_text.startswith("Home /"): continue

            line_content = ""
            if element.name == 'li': line_content = "• " 

            for child in element.children:
                child_text = ""
                if isinstance(child, NavigableString): child_text = str(child)
                else: child_text = child.get_text()
                
                if child.name in ['strong', 'b']:
                    line_content += f"**{child_text.strip()}** "
                else:
                    line_content += child_text
            
            clean_line = line_content.strip()
            if any(x in clean_line.lower() for x in stop_keywords): break
            if "about the journal" in clean_line.lower() and len(clean_line) < 30: continue

            if clean_line: full_text_lines.append(clean_line)
            
    return "\n\n".join(full_text_lines)

def generate_from_template(template_file, cover_image_file, journal_details, role_data_dict, about_data, toc_sections, list_vars):
    """
    Core document generation engine. Maps extracted metadata to the MS Word template payload.
    """
    doc = DocxTemplate(template_file)
    
    if cover_image_file:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                tmp.write(cover_image_file.getvalue())
                tmp_path = tmp.name
            doc.replace_pic('CoverPlaceholder', tmp_path)
            os.remove(tmp_path)
        except Exception as e:
            print(f"Warning: Could not swap cover image. Error: {e}")

    def get_clean_list(text_block):
        if not text_block: return []
        return [line.strip() for line in str(text_block).split('\n') if line.strip()]

    # Smart Journal Name Splitter
    raw_journal_name = journal_details.get('journal_name', '')
    
    # Looks for a string ending with parentheses, e.g., "Full Name (SHORT)"
    match = re.match(r'^(.*?)\s*\(([^)]+)\)$', raw_journal_name)
    if match:
        clean_name = match.group(1).strip()
        short_name = match.group(2).strip()
    else:
        clean_name = raw_journal_name
        short_name = ""

    # Load the standard system variables
    context = {
        'journal_name': clean_name,         
        'journal_short': short_name,        
        'journal_full': raw_journal_name,   
        
        'vol': journal_details.get('volume', ''),
        'issue': journal_details.get('issue', ''),
        'month': journal_details.get('month', ''),
        'year': journal_details.get('year', ''),
        'eissn': journal_details.get('eissn', ''),
        
        # Pass the 'doc' object here so it can build the hyperlinks
        'about_text': make_rich_text(about_data.get('about', ''), doc),
        'address_text': make_rich_text(about_data.get('address', ''), doc),
        'toc_sections': toc_sections
    }
    
    # Dynamically load ALL detected template roles into the context
    for role_var, text_val in role_data_dict.items():
        if role_var in list_vars:
            context[role_var] = get_clean_list(text_val) 
        else:
            context[role_var] = make_rich_text(text_val, doc)
            
    doc.render(context)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def fix_toc_sorting(df):
    """
    Sorts the Table of Contents via DOI suffixes and prioritizes Editorial Previews.
    """
    doi_col = next((c for c in df.columns if c.lower() == 'doi'), None)
    title_col = next((c for c in df.columns if 'title' in c.lower() and 'section' not in c.lower()), None)
    page_col = 'Pages_Clean' if 'Pages_Clean' in df.columns else next((c for c in df.columns if 'page' in c.lower()), None)

    if not doi_col or not title_col: return df 

    df['temp_doi_suffix'] = df[doi_col].astype(str).str.extract(r'(\d+)$').astype(float)

    mask = df[title_col].astype(str).str.contains("Editorial Preview", case=False, na=False)
    df.loc[mask, 'temp_doi_suffix'] = -1
    if page_col:
        df.loc[mask, page_col] = 'i'

    if 'Section Name' in df.columns:
        def get_rank(name):
            n = str(name).lower()
            if any(k in n for k in ['regular', 'article', 'original']): return 0
            return 1

        df['section_rank'] = df['Section Name'].apply(get_rank)
        df = df.sort_values(by=['section_rank', 'Section Name', 'temp_doi_suffix'], ascending=[True, True, True])
        df = df.drop(columns=['section_rank'])
    else:
        df = df.sort_values(by=['temp_doi_suffix'], ascending=True)

    df = df.drop(columns=['temp_doi_suffix'])
    return df

def fetch_pages_from_crossref(doi_string):
    """
    Pings the public Crossref API to retrieve page numbers for a given DOI.
    """
    if not doi_string or str(doi_string).lower() == 'nan': 
        return ""
        
    # Clean the DOI in case the CSV exported it as a full URL
    clean_doi = str(doi_string).replace('https://doi.org/', '').strip()
    
    try:
        # Ping the Crossref REST API
        url = f"https://api.crossref.org/works/{clean_doi}"
        response = requests.get(url, timeout=5) # 5 second timeout so app doesn't freeze
        
        if response.status_code == 200:
            data = response.json()
            return data.get('message', {}).get('page', '')
    except Exception as e:
        print(f"Crossref API Error for {clean_doi}: {e}")
        
    return ""

def extract_submission_id(filename):
    """
    Extracts the OJS submission ID from both new and legacy filename formats.
    """
    clean_filename = urllib.parse.unquote(filename).lower()
    
    # Check for the NEW format: ID is at the start, followed by a hyphen OR an underscore
    new_format_match = re.search(r'^(\d+)[-_]', clean_filename)
    if new_format_match:
        return new_format_match.group(1)
        
    # Check for the OLD format: ID is preceded by a '+', followed by a hyphen OR an underscore
    old_format_match = re.search(r'\+(\d+)[-_]', clean_filename)
    if old_format_match:
        return old_format_match.group(1)
        
    return None

# ==========================================
#        PART 2: MAIN APP UI
# ==========================================

st.set_page_config(page_title="Pre-Pages Generator", layout="wide")
st.title("Pre-Pages Generator & Visualiser")

tab_gen, tab_viz = st.tabs(["📄 Document Generator", "📊 Data Visualiser"])

# --- SESSION STATE INITIALIZATION ---
if 'step' not in st.session_state: st.session_state.step = 1
if 'scraped_data' not in st.session_state: st.session_state.scraped_data = None
if 'scraped_about' not in st.session_state: st.session_state.scraped_about = ""
if 'raw_toc_df' not in st.session_state: st.session_state.raw_toc_df = None 
if 'last_uploaded_file' not in st.session_state: st.session_state.last_uploaded_file = None
if 'editor_df' not in st.session_state: st.session_state.editor_df = None   
if 'final_doc_buffer' not in st.session_state: st.session_state.final_doc_buffer = None

current_year_int = datetime.date.today().year
year_list = [str(y) for y in range(current_year_int - 5, current_year_int + 6)]
default_year_idx = 5 

if 'inputs' not in st.session_state:
    st.session_state.inputs = {
        'journal_name': '', 'volume': '', 'issue': '', 
        'month': 'January', 'year': str(current_year_int), 'eissn': '',
        'about_text': '', 'address_text': DEFAULT_ADDRESS,
        'eic_val': '', 'me_val': '', 'co_val': '', 'se_val': '', 'ab_val': ''
    }

def next_step(): st.session_state.step += 1
def prev_step(): st.session_state.step -= 1

# ==========================================
#        TAB 1: DOCUMENT GENERATOR
# ==========================================
with tab_gen:
    
    # ------------------------------------------
    # STEP 1: SETUP & TEMPLATE PARSING
    # ------------------------------------------
    if st.session_state.step == 1:
        st.header("Step 1: Setup")
        st.info("ℹ️ Important: In your Word Template, right-click the dummy cover image > View Alt Text > enter 'CoverPlaceholder'.")
        
        c1, c2 = st.columns(2)
        uploaded_template = c1.file_uploader("1. Word Template (.docx)", type=['docx'], key='u_t')
        uploaded_cover = c2.file_uploader("2. Cover Image (JPG/PNG)", type=['jpg', 'png', 'jpeg'], key='u_c')
        
        if uploaded_template: 
            st.session_state.inputs['template_file'] = uploaded_template
            
            # --- Extract Journal Short Form from Filename (e.g., "JIWE_template.docx" -> "JIWE") ---
            filename = uploaded_template.name
            if "_" in filename:
                st.session_state.detected_journal_short = filename.split("_")[0].upper()
            else:
                st.session_state.detected_journal_short = None

            # --- Analyze OOXML Template structure for custom Jinja2 tags ---
            roles, list_vars = analyze_template(uploaded_template)
            st.session_state.template_roles = roles
            st.session_state.template_list_vars = list_vars
            
            c1.success(f"📄 Template parsed! Found {len(roles)} custom editorial roles.")
            
        if uploaded_cover: 
            st.session_state.inputs['cover_image'] = uploaded_cover
            c2.image(uploaded_cover, caption="Cover Image Preview", use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("Next >"):
            if 'template_file' in st.session_state.inputs: 
                next_step()
                st.rerun()
            else: 
                st.error("Please upload a Word Template to proceed.")

    # ------------------------------------------
    # STEP 2: JOURNAL DETAILS (BaaS INTEGRATION)
    # ------------------------------------------
    elif st.session_state.step == 2:
        st.header("Step 2: Issue Details")

        # Sync UI widget states BEFORE calculating dropdown indexes
        if 'ui_journal' in st.session_state: st.session_state.inputs['journal_name'] = st.session_state.ui_journal
        if 'ui_vol' in st.session_state: st.session_state.inputs['volume'] = st.session_state.ui_vol
        if 'ui_issue' in st.session_state: st.session_state.inputs['issue'] = st.session_state.ui_issue
        if 'ui_month' in st.session_state: st.session_state.inputs['month'] = st.session_state.ui_month
        if 'ui_year' in st.session_state: st.session_state.inputs['year'] = st.session_state.ui_year

        # --- SUPABASE FETCH (Journals, eISSN, Address) ---
        db_journals_data = []
        if supabase:
            try:
                response = supabase.table("journals").select("id, name, eissn, address").execute()
                db_journals_data = response.data
            except Exception as e:
                st.error(f"Database error: {e}")
        
        # Fallback if DB is empty or disconnected
        if not db_journals_data:
            db_journals_data = [{"id": 0, "name": "Journal of Informatics and Web Engineering (JIWE)", "eissn": "2821-3636", "address": "Faculty of Computing and Informatics, Multimedia University, 63100 Cyberjaya, Selangor, Malaysia."}]
            
        journal_names = [row['name'] for row in db_journals_data]
        journal_map = {row['name']: row.get('eissn', '') or '' for row in db_journals_data}
        address_map = {row['name']: row.get('address', '') or '' for row in db_journals_data}
        journal_opts = journal_names + ["➕ Add New Journal..."]
        
        # Smart Auto-Select Logic (Matches uploaded template filename to DB records)
        saved_journal = st.session_state.inputs.get('journal_name', '')
        detected_short = st.session_state.get('detected_journal_short')
        j_idx = 0 
        
        if saved_journal in journal_opts and saved_journal != '':
            j_idx = journal_opts.index(saved_journal)
        elif detected_short:
            for i, name in enumerate(journal_names):
                if f"({detected_short})" in name:
                    j_idx = i
                    st.session_state.inputs['journal_name'] = name 
                    break
        
        # --- UI: DROPDOWN & DELETE BUTTON ---
        col_j1, col_j2 = st.columns([5, 1])
        with col_j1:
            selected_journal = st.selectbox("Journal Name", journal_opts, index=j_idx, key="ui_journal")
        with col_j2:
            st.markdown("<br>", unsafe_allow_html=True) 
            if selected_journal != "➕ Add New Journal...":
                if st.button("🗑️ Delete", help="Remove this journal from the database"):
                    if supabase:
                        try:
                            supabase.table("journals").delete().eq("name", selected_journal).execute()
                            st.success(f"Deleted '{selected_journal}'")
                            
                            st.session_state.inputs['journal_name'] = journal_names[0] if len(journal_names) > 1 else ""
                            if 'ui_journal' in st.session_state: del st.session_state['ui_journal'] 
                            st.rerun()
                        except Exception as e:
                            st.error(f"Failed to delete: {e}")

        # --- UI: ADD NEW JOURNAL ---
        if selected_journal == "➕ Add New Journal...":
            st.info("💡 Add a new journal to the database. It will be saved permanently.")
            col_n1, col_n2 = st.columns([3, 1])
            new_journal = col_n1.text_input("Enter new Journal Name:")
            new_eissn = col_n2.text_input("Enter eISSN (Optional):")
            new_address = st.text_area("Enter Publisher Address (Optional):", height=100)
            
            if st.button("Save New Journal", type="secondary"):
                if new_journal and new_journal not in journal_names:
                    if supabase:
                        try:
                            supabase.table("journals").insert({
                                "name": new_journal, 
                                "eissn": new_eissn, 
                                "address": new_address
                            }).execute()
                            st.success(f"Added '{new_journal}' to the database!")
                        except Exception as e:
                            st.error(f"Failed to save to database: {e}")
                            
                    st.session_state.inputs['journal_name'] = new_journal
                    if 'ui_journal' in st.session_state: del st.session_state['ui_journal'] 
                    st.rerun() 
                elif new_journal in journal_names:
                    st.warning("Journal already exists in the list.")
        else:
            st.session_state.inputs['journal_name'] = selected_journal

            # Smart Address Auto-Fill: Updates address only when journal selection changes
            if st.session_state.get('last_selected_journal') != selected_journal:
                st.session_state.inputs['address_text'] = address_map.get(selected_journal, '')
                st.session_state['last_selected_journal'] = selected_journal

        # --- VOLUME AND ISSUE LOGIC ---
        c1, c2 = st.columns(2)
        vol_opts = [str(i) for i in range(1, 101)]
        issue_opts = [str(i) for i in range(1, 51)]
        
        saved_vol = str(st.session_state.inputs.get('volume', '5'))
        saved_issue = str(st.session_state.inputs.get('issue', '1'))
        v_idx = vol_opts.index(saved_vol) if saved_vol in vol_opts else 4
        i_idx = issue_opts.index(saved_issue) if saved_issue in issue_opts else 0
        
        c1.selectbox("Volume", vol_opts, index=v_idx, key="ui_vol")
        c2.selectbox("Issue", issue_opts, index=i_idx, key="ui_issue")
        
        # --- MONTH/YEAR & AUTO-FILL eISSN LOGIC ---
        c3, c4, c5 = st.columns([2, 2, 2])
        month_opts = ["January","February","March","April","May","June","July","August","September","October","November","December"]
        saved_month = st.session_state.inputs.get('month', 'January')
        saved_year = st.session_state.inputs.get('year', str(current_year_int))
        
        try: m_idx = month_opts.index(saved_month)
        except ValueError: m_idx = 0
        try: y_idx = year_list.index(saved_year)
        except ValueError: y_idx = default_year_idx

        c3.selectbox("Month", month_opts, index=m_idx, key="ui_month")
        c4.selectbox("Year", year_list, index=y_idx, key="ui_year")
        
        default_eissn = journal_map.get(selected_journal, '') if selected_journal != "➕ Add New Journal..." else ''
        st.session_state.inputs['eissn'] = c5.text_input("eISSN", value=default_eissn)
        
        st.markdown("<br>", unsafe_allow_html=True)
        col_b, col_n = st.columns([1, 6])
        with col_b: st.button("< Back", on_click=prev_step, key="btn_back_2")
        with col_n: 
            if selected_journal != "➕ Add New Journal...":
                if st.button("Next >", key="btn_next_2"): next_step(); st.rerun()
            else:
                st.button("Next >", disabled=True, help="Please save the new journal name before proceeding.")

    # ------------------------------------------
    # STEP 3: EDITORIAL BOARD (WEB SCRAPER)
    # ------------------------------------------
    elif st.session_state.step == 3:
        st.header("Step 3: Editorial Board Mapper")
        st.write("The fields below are dynamically generated based on the tags inside the Word Template.")
        
        url_input = st.text_input("Website URL to Scrape:", placeholder="https://journals.mmupress.com/index.php/jiwe/about/editorialTeam")
        
        if st.button("🔍 Fetch Live Roles"):
            with st.spinner("Analyzing website structure..."):
                try:
                    headers = {'User-Agent': 'Mozilla/5.0'}
                    response = requests.get(url_input, headers=headers)
                    soup = BeautifulSoup(response.content, 'html.parser')
                    st.session_state['scraped_data'] = parse_editorial_html(soup)
                    st.success("Successfully scraped dynamic roles!")
                except Exception as e: 
                    st.error(f"Error connecting to website: {e}")
        
        scraped_data = st.session_state.get('scraped_data') or {}
        scraped_options = ["-- Leave Blank / Manual Entry --"] + list(scraped_data.keys())
        
        st.markdown("### Map Extracted Data to Template Fields")
        
        # Pull the custom roles detected from the Word Template in Step 1
        roles_to_map = st.session_state.get('template_roles', [])
        
        if not roles_to_map:
            st.warning("⚠️ No custom roles found in the template. Did you use tags like `{{ editor_in_chief }}`?")
            
        # Define the mapper function INSIDE Step 3
        def role_mapper(role_var):
            # Format "editor_in_chief" to "Editor In Chief" for the UI
            display_label = role_var.replace('_', ' ').title()
            st.markdown(f"**{display_label}**")

            # 1. Define a callback function to update the text_area when selectbox changes
            def update_text():
                new_source = st.session_state[f"src_{role_var}"]
                if new_source != "-- Leave Blank / Manual Entry --":
                    st.session_state[f"role_val_{role_var}"] = "\n".join(scraped_data[new_source])

            col_src, col_txt = st.columns([1, 2])

            with col_src:
                # Auto-match logic
                default_idx = 0
                for i, opt in enumerate(scraped_options):
                    if opt.lower().replace('-', ' ') == display_label.lower().replace('-', ' '):
                        default_idx = i
                        break
                
                # Use on_change to trigger the update instantly
                st.selectbox(
                    f"Source for {role_var}", 
                    scraped_options, 
                    index=default_idx, 
                    key=f"src_{role_var}",
                    on_change=update_text,
                    label_visibility="collapsed"
                )

            with col_txt:
                # If the value hasn't been set yet, initialize it
                if f"role_val_{role_var}" not in st.session_state:
                    st.session_state[f"role_val_{role_var}"] = st.session_state.inputs.get(f"role_{role_var}", '')

                # The text area now listens to the session state variable
                final_text = st.text_area(
                    f"text for {role_var}", 
                    key=f"role_val_{role_var}", 
                    height=150, 
                    label_visibility="collapsed"
                )
                st.session_state.inputs[f"role_{role_var}"] = final_text
                
        # Generate the UI for EXACTLY the roles found in the template
        for r_var in roles_to_map:
            role_mapper(r_var)
        
        st.markdown("<br>", unsafe_allow_html=True)
        col_b, col_n = st.columns([1, 6])
        with col_b: st.button("< Back", on_click=prev_step, key="btn_back_3")
        with col_n: 
            if st.button("Next >", key="btn_next_3"): next_step(); st.rerun()

    # ------------------------------------------
    # STEP 4: ABOUT THE JOURNAL (SANITIZER)
    # ------------------------------------------
    elif st.session_state.step == 4:
        st.header("Step 4: About")
        url = st.text_input("Fetch URL")
        if st.button("Fetch"):
            try:
                headers = {'User-Agent': 'Mozilla/5.0'}
                response = requests.get(url, headers=headers)
                soup = BeautifulSoup(response.content, 'html.parser')
                
                # Wrap the scraped text in the sanitizer to prevent Word layout issues
                raw_about = parse_about_html(soup)
                st.session_state['scraped_about'] = clean_scraped_text(raw_about)
                
                st.success("Fetched and Sanitized!")
            except: 
                st.error("Fetch Error")
            
        st.session_state.inputs['about_text'] = st.text_area("About", value=st.session_state.get('scraped_about', '') or st.session_state.inputs.get('about_text', ''), height=250)
        
        # Address auto-populates based on Step 2 DB fetch
        st.session_state.inputs['address_text'] = st.text_area("Address", value=st.session_state.inputs.get('address_text', ''), height=150)
        
        col_b, col_n = st.columns([1, 6])
        with col_b: st.button("< Back", on_click=prev_step, key="btn_back_4")
        with col_n: 
            if st.button("Next >", key="btn_next_4"): next_step(); st.rerun()

    # ------------------------------------------
    # STEP 5: TABLE OF CONTENTS & DOCUMENT GEN
    # ------------------------------------------
    elif st.session_state.step == 5:
        st.header("Step 5: Table of Contents")
        c1, c2 = st.columns(2)
        uploaded_file = c1.file_uploader("1. Upload CSV", type=['csv'], key='toc_uploader')
        uploaded_pdfs = c2.file_uploader("2. Upload PDFs or ZIP (Optional)", type=['pdf', 'zip'], accept_multiple_files=True, key='pdf_uploader')
        
        if uploaded_file:
            try:
                if st.session_state.raw_toc_df is None or st.session_state.last_uploaded_file != uploaded_file.name:
                    temp_df = clean_ojs_csv(pd.read_csv(uploaded_file))
                    
                    # --- SECURITY FAILSAFE (Prevents DB Corruption) ---
                    csv_shortcode = extract_shortcode_from_csv(temp_df)
                    ui_journal_name = st.session_state.inputs.get('journal_name', '')
                    
                    ui_match = re.search(r'\(([^)]+)\)$', ui_journal_name)
                    ui_shortcode = ui_match.group(1).lower() if ui_match else ""
                    
                    if csv_shortcode and ui_shortcode and csv_shortcode != ui_shortcode:
                        st.error(f"🚨 **Security Alert:** You selected **{ui_journal_name}** in Step 2, but uploaded a CSV belonging to **{csv_shortcode.upper()}**.")
                        st.warning("Upload canceled to prevent database corruption. Please check your files and try again.")
                        st.stop()

                    st.session_state.raw_toc_df = fix_toc_sorting(temp_df)
                    st.session_state.last_uploaded_file = uploaded_file.name
                    st.session_state.editor_df = None 

                    # --- SUPABASE DATA SYNC ---
                    if supabase:
                        with st.spinner("Syncing data to Cloud Database..."):
                            db_records = []
                            full_df = st.session_state.raw_toc_df
                            
                            t_col = next((c for c in full_df.columns if 'title' in c.lower() and 'section' not in c.lower()), None)
                            a_col = 'Authors_Combined' if 'Authors_Combined' in full_df.columns else None
                            s_col = 'Section Name' if 'Section Name' in full_df.columns else None
                            st_col = next((c for c in full_df.columns if 'status' in c.lower()), None)

                            final_shortcode = ui_shortcode if ui_shortcode else csv_shortcode

                            for _, row in full_df.iterrows():
                                db_records.append({
                                    "journal_name": ui_journal_name,
                                    "journal_short": final_shortcode,
                                    "article_title": str(row[t_col]) if t_col else "Unknown",
                                    "authors": str(row[a_col]) if a_col else "Unknown",
                                    "section_name": str(row[s_col]) if s_col else "Articles",
                                    "status": str(row[st_col]) if st_col else "Unknown"
                                })
                            
                            try:
                                supabase.table("issue_metadata").delete().eq("journal_name", ui_journal_name).execute()
                                supabase.table("issue_metadata").insert(db_records).execute()
                            except Exception as e:
                                st.error(f"Failed to sync with database: {e}")

                full_df = st.session_state.raw_toc_df
                
                # --- STATUS FILTERING ---
                status_col = next((c for c in full_df.columns if 'status' in c.lower()), None)
                filtered_df_for_doc = full_df
                
                if status_col:
                    unique_statuses = full_df[status_col].unique().tolist()
                    st.subheader("Filter Articles (For Document Only)")
                    default_opts = [x for x in unique_statuses if 'scheduled' in str(x).lower()] or unique_statuses
                    selected_status = st.multiselect("Select Statuses:", options=unique_statuses, default=default_opts)
                    filtered_df_for_doc = full_df[full_df[status_col].isin(selected_status)]

                # --- BUILD THE EDITABLE DATAFRAME ---
                if st.session_state.editor_df is None:
                    cols = filtered_df_for_doc.columns.tolist()
                    def find_col(keywords): return next((c for c in cols if c.lower() in keywords), None)
                    rename_map = {}
                    if c := find_col(['submission id', 'id']): rename_map[c] = "Submission ID"
                    if c := find_col(['article title', 'title']): rename_map[c] = "Article Title"
                    if c := find_col(['authors', 'authors_combined']): rename_map[c] = "Authors"
                    if c := find_col(['page numbers', 'pages_clean', 'pages']): rename_map[c] = "Page Numbers"
                    if c := find_col(['section name', 'section']): rename_map[c] = "Section Name"
                    if c := find_col(['doi']): rename_map[c] = "DOI" 
                    
                    display_df = filtered_df_for_doc.rename(columns=rename_map)
                    
                    for req in ["Submission ID", "Section Name", "Article Title", "Authors", "DOI", "Page Numbers"]:
                        if req not in display_df.columns: display_df[req] = "Articles" if req == "Section Name" else ""
                        
                    df_final_cols = display_df[["Submission ID", "Section Name", "Article Title", "Authors", "DOI", "Page Numbers"]].copy()
                    
                    # Inject Order column as string to preserve '00' logic
                    df_final_cols.insert(0, 'Order', [str(i) for i in range(1, 1 + len(df_final_cols))])
                    
                    st.session_state.editor_df = df_final_cols

                st.subheader("Edit Content (Document Version)")

                st.info("""
                **💡 Pro-Tip: Advanced Sorting & Formatting**
                Use the **Order** column on the far left to easily reorganize your Table of Contents:
                * Type **0 (zero)** to move an article to the very **top of its current section**.
                * Type **00 (double zero)** to move an article to the top, push its **entire section to the very top**, AND automatically apply **Roman numeral page numbers (i, ii, iii)**!
                *(Highly useful for positioning Editorial Previews and Front Matter!)*
                """)
                
                col_api, col_pdf = st.columns(2)
                
                # --- HYBRID PAGE ENGINE: CROSSREF API ---
                with col_api:
                    if st.button("🌐 Auto-Fetch Pages via DOI"):
                        with st.spinner("Pinging Crossref database..."):
                            updated_df = st.session_state.editor_df.copy()
                            match_count = 0
                            for idx, row in updated_df.iterrows():
                                doi_val = row.get('DOI', '')
                                if doi_val and str(doi_val).strip():
                                    fetched_pages = fetch_pages_from_crossref(doi_val)
                                    if fetched_pages:
                                        updated_df.at[idx, 'Page Numbers'] = fetched_pages
                                        match_count += 1
                            st.session_state.editor_df = updated_df
                            if match_count > 0:
                                st.success(f"Fetched pages for {match_count} articles from Crossref!")
                            else:
                                st.warning("No page numbers found in Crossref for these DOIs. Try the PDF method.")
                            st.rerun()

                # --- HYBRID PAGE ENGINE: PDF EXTRACTION ---
                with col_pdf:
                    if uploaded_pdfs and not st.session_state.editor_df.empty:
                        if st.button("✨ Auto-Calculate Pages from PDFs"):
                            pdf_map = {}
                            for file in uploaded_pdfs:
                                if file.name.lower().endswith('.zip'):
                                    with zipfile.ZipFile(file, 'r') as z:
                                        for z_name in z.namelist():
                                            if z_name.lower().endswith('.pdf') and not z_name.startswith('__MACOSX') and not z_name.startswith('.'):
                                                try:
                                                    pdf_bytes = BytesIO(z.read(z_name))
                                                    clean_fname = os.path.basename(z_name).lower().replace('.pdf', '').strip()
                                                    pdf_map[clean_fname] = len(PdfReader(pdf_bytes).pages)
                                                except Exception as e:
                                                    print(f"Error reading {z_name} from ZIP: {e}")
                                elif file.name.lower().endswith('.pdf'):
                                    try: 
                                        clean_fname = file.name.lower().replace('.pdf', '').strip()
                                        pdf_map[clean_fname] = len(PdfReader(file).pages)
                                    except: 
                                        pass
                                        
                            current_roman_page = 1
                            current_arabic_page = 1
                            updated_df = st.session_state.editor_df.copy()
                            matched_count = 0
                            
                            for idx, row in updated_df.iterrows():
                                title_clean = str(row['Article Title']).lower()
                                
                                # Sanitize ID: Remove '.0' if pandas upcast the column to float due to empty rows in the raw CSV
                                sub_id = str(row.get('Submission ID', '')).replace('.0', '').strip()
                                page_count = 0
                                
                                for fname, count in pdf_map.items():
                                    # Use strict regex extractor
                                    extracted_id = extract_submission_id(fname)
                                    
                                    # Strict '==' match prevents false positives (e.g., ID 15 matching file 1583)
                                    is_id_match = (sub_id != "" and sub_id != "nan" and sub_id == extracted_id)
                                    is_title_match = (fname in title_clean or title_clean[:30] in fname)
                                    
                                    if is_id_match or is_title_match:
                                        page_count = count
                                        matched_count += 1
                                        break
                                        
                                if page_count > 0:
                                    order_val = str(updated_df.at[idx, 'Order']).strip()
                                    
                                    # Apply Roman numerals only to "00" priority front-matter
                                    if order_val == "00":
                                        start_str = int_to_roman(current_roman_page)
                                        updated_df.at[idx, 'Page Numbers'] = start_str
                                        current_roman_page += page_count
                                    else:
                                        updated_df.at[idx, 'Page Numbers'] = str(current_arabic_page)
                                        current_arabic_page += page_count
                                        
                            st.session_state.editor_df = updated_df
                            st.success(f"Calculated starting pages for {matched_count} articles!")
                            st.rerun()

                edited_df = st.data_editor(st.session_state.editor_df, use_container_width=True, num_rows="dynamic", key="data_editor_box")

                st.info("💡 **Important Note for PDF Page Calculation:** To successfully use the **'Auto-Calculate Pages from PDFs'** feature, please ensure that your uploaded PDF filenames begin with their respective Open Journal Systems (OJS) Submission ID (e.g., *`3576-Article_Text.pdf`*). The system relies on this ID to accurately match the physical file to the correct row in the Table of Contents above.")

            except Exception as e: st.error(f"Error: {e}")

        st.markdown("---")
        col1, col2, col3 = st.columns([1, 6, 3])
        with col1: st.button("< Back", on_click=prev_step, key="btn_back_5")
        
        # --- FINAL DOCUMENT GENERATION ---
        with col3:
            if st.button("Generate Document", type="primary"):
                
                # 1. Identify Priority Sections (The "00" Logic)
                priority_sections = []
                for idx, row in edited_df.iterrows():
                    val = str(row.get('Order', '')).strip() 
                    if val == '00':
                        sec_name = row['Section Name']
                        if sec_name not in priority_sections:
                            priority_sections.append(sec_name)

                # 2. Convert Order to Numeric for sorting Articles
                edited_df['Order_Num'] = pd.to_numeric(edited_df['Order'], errors='coerce').fillna(999)

                # 3. Determine the Master Section Order
                original_sections = []
                if "Section Name" in st.session_state.raw_toc_df.columns:
                    original_sections = pd.unique(st.session_state.raw_toc_df["Section Name"]).tolist()
                
                final_section_order = priority_sections.copy()
                for sec in original_sections:
                    if sec not in final_section_order and sec in edited_df["Section Name"].values:
                        final_section_order.append(sec)
                
                # Catch any stray sections
                for sec in edited_df["Section Name"].unique():
                    if sec not in final_section_order:
                        final_section_order.append(sec)

                # 4. Package Data for Jinja2 Template
                toc_sections_ready = []
                for section_name in final_section_order:
                    sec_df = edited_df[edited_df["Section Name"] == section_name]
                    
                    if not sec_df.empty:
                        sec_df = sec_df.sort_values(by="Order_Num")
                        
                        articles = []
                        for index, row in sec_df.iterrows():
                            articles.append({
                                'title': str(row['Article Title']),
                                'authors': str(row['Authors']),
                                'pages': str(row['Page Numbers'])
                            })
                        toc_sections_ready.append({'name': section_name, 'articles': articles})

                # --- Execute Engine ---
                j_details = st.session_state.inputs
                about_data = {'about': st.session_state.inputs['about_text'], 'address': st.session_state.inputs['address_text']}

                role_data_dict = {}
                roles_to_map = st.session_state.get('template_roles', [])
                for r_var in roles_to_map:
                    role_data_dict[r_var] = st.session_state.inputs.get(f"role_{r_var}", "")
                
                list_vars = st.session_state.get('template_list_vars', [])

                try:
                    final_docx_buffer = generate_from_template(
                        st.session_state.inputs['template_file'], 
                        st.session_state.inputs.get('cover_image'), 
                        j_details, 
                        role_data_dict, 
                        about_data, 
                        toc_sections_ready,
                        list_vars 
                    )
                    st.success("🎉 Document generated successfully!")
                    st.download_button("Download Final Doc (.docx)", final_docx_buffer, "pre_pages_final.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e: 
                    st.error(f"Template Error: {e}")
            

# ==========================================
#        TAB 2: DATA VISUALISER
# ==========================================
with tab_viz:
    st.header("📊 Editorial Intelligence Dashboard")
    st.write("Visualizations are based on **Active Papers Only** (Published, Scheduled, Production, Copyediting).")

    db_metadata = None
    
    # --- BaaS DATABASE FETCH ---
    if supabase:
        try:
            # 1. Fetch a list of all unique journals currently in the database
            journals_response = supabase.table("issue_metadata").select("journal_name").execute()
            
            if journals_response.data:
                # Get unique names and filter out any None/Blank values
                active_journals = list(set([row['journal_name'] for row in journals_response.data if row.get('journal_name')]))
                
                if active_journals:
                    st.markdown("### Select Journal to Analyze")
                    
                    # Global "All Journals" option
                    dropdown_options = ["🌐 All Journals (Merged Data)"] + sorted(active_journals)
                    
                    # Default to the journal being worked on in Tab 1
                    current_tab1_journal = st.session_state.inputs.get('journal_name', '')
                    default_index = dropdown_options.index(current_tab1_journal) if current_tab1_journal in dropdown_options else 0
                    
                    viz_journal_selected = st.selectbox("Journal Filter:", dropdown_options, index=default_index, key="viz_journal_select")
                    
                    # Global vs. Specific Journal routing
                    if viz_journal_selected == "🌐 All Journals (Merged Data)":
                        # Fetch ALL data across the entire MMU Press database
                        response = supabase.table("issue_metadata").select("*").execute()
                    else:
                        # Fetch data ONLY for the selected journal
                        response = supabase.table("issue_metadata").select("*").eq("journal_name", viz_journal_selected).execute()
                    
                    if response.data:
                        db_metadata = pd.DataFrame(response.data)
                        db_metadata = db_metadata.rename(columns={
                            "article_title": "Article Title",
                            "authors": "Authors_Combined",
                            "section_name": "Section Name",
                            "status": "status"
                        })
        except Exception as e:
            st.error(f"Failed to fetch visualization data from database: {e}")

    # Fallback to local session state if DB is empty or fails
    df_source = db_metadata if db_metadata is not None and not db_metadata.empty else st.session_state.raw_toc_df

    if df_source is not None and not df_source.empty:
        raw_df = df_source.copy()
        
        # --- DATA FILTERING (Active Status) ---
        df_viz_source = raw_df 
        status_col = next((c for c in raw_df.columns if 'status' in c.lower()), None)
        
        if status_col:
            active_keywords = ['published', 'scheduled', 'production', 'copyediting']
            mask = raw_df[status_col].astype(str).apply(lambda x: any(k in x.lower() for k in active_keywords))
            df_active = raw_df[mask]
            
            if df_active.empty:
                st.warning("⚠️ No active papers found. Visualizing ALL data instead.")
                df_viz_source = raw_df
            else:
                st.success(f"Filtering applied: Visualizing {len(df_active)} active papers out of {len(raw_df)} total submissions.")
                df_viz_source = df_active
        
        # Standardize columns for visualization modules
        cols = df_viz_source.columns
        def get_c(k):
            for c in cols:
                if any(s in c.lower() for s in k): return c
            return None
        
        viz_map = {}
        if c := get_c(['title', 'article']): viz_map[c] = "Article Title"
        if 'Authors_Combined' in df_viz_source.columns: viz_map['Authors_Combined'] = "Authors"
        if 'Section Name' in df_viz_source.columns: viz_map['Section Name'] = "Section Name"
        
        df_viz = df_viz_source.rename(columns=viz_map)
        
        st.markdown("---")
        
        # ==========================================
        # UPGRADE 1: THE EXECUTIVE KPI ROW
        # ==========================================
        st.markdown("### 📈 Executive Summary")
        
        total_subs = len(raw_df)
        active_count = len(df_viz)
        
        unique_authors = 0
        if 'Authors' in df_viz.columns:
            all_auths = df_viz['Authors'].dropna().str.split(',').explode().str.strip()
            unique_authors = all_auths.nunique()
            
        top_section = "N/A"
        if 'Section Name' in df_viz.columns and not df_viz.empty:
            top_section = df_viz['Section Name'].mode()[0]

        # Custom CSS Injection: Aggressively center every sub-element of the metric cards
        st.markdown("""
        <style>
        /* Center the main metric container */
        div[data-testid="stMetric"] {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
        }
        /* Force the Title to center */
        div[data-testid="stMetricLabel"] {
            display: flex;
            justify-content: center;
            width: 100%;
        }
        /* Force the Number to center */
        div[data-testid="stMetricValue"] {
            display: flex;
            justify-content: center;
            width: 100%;
        }
        /* Force the Delta (Conversion Rate) to center */
        div[data-testid="stMetricDelta"] {
            display: flex;
            justify-content: center;
            width: 100%;
        }
        </style>
        """, unsafe_allow_html=True)
            
        # Grouped KPI bounding box
        with st.container(border=True):
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total Submissions", total_subs)
            k2.metric("Active Papers", active_count, delta=f"{round((active_count/total_subs)*100)}% Conversion" if total_subs > 0 else None)
            k3.metric("Unique Authors", unique_authors)
            k4.metric("Top Section", top_section)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # ==========================================
        # UPGRADE 2: THE VISUALIZATION GRID
        # ==========================================
        col1, col2 = st.columns(2)
        
        with col1:
            fig_sec = utils_viz.plot_sections_pie(df_viz)
            if fig_sec: 
                with st.container(border=True):
                    st.plotly_chart(fig_sec, use_container_width=True)
            else: 
                st.warning("Section information missing.")
            
        with col2:
            fig_auth = utils_viz.plot_top_authors(df_viz)
            if fig_auth: 
                with st.container(border=True):
                    st.plotly_chart(fig_auth, use_container_width=True)
            else: 
                st.warning("Author information missing.")

        st.markdown("<br>", unsafe_allow_html=True)
        
        # ==========================================
        # UPGRADE 3: TOPIC TRENDS (WORD CLOUD)
        # ==========================================
        st.markdown("### 🧠 Topic Trends (Word Cloud)")
        with st.container(border=True):
            wc_image = utils_viz.generate_wordcloud(df_viz)
            
            if wc_image is not None: 
                # Convert raw pixel array to PNG binary data
                pil_img = Image.fromarray(wc_image)
                img_buffer = io.BytesIO()
                pil_img.save(img_buffer, format="PNG")
                b64_string = base64.b64encode(img_buffer.getvalue()).decode()
                
                # Render via native HTML to completely bypass the Streamlit micro-dot resizing bug
                st.markdown(
                    f'<img src="data:image/png;base64,{b64_string}" style="width: 100%; height: auto; display: block; border-radius: 4px;">', 
                    unsafe_allow_html=True
                )
            else: 
                st.warning("Article Titles required for Word Cloud.")

    else:
        st.info("👋 Welcome! It looks like there is no active issue in the database. Please generate an issue in the Generator tab to begin tracking analytics.")
