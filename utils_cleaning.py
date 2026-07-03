import pandas as pd

def clean_ojs_csv(df):
    """
    CLEANER FOR OJS CSV FILES
    Standardizes column names, merges author columns (up to 15), 
    and normalizes page numbers and section names.
    """
    # 1. NORMALIZE COLUMN NAMES
    # Strip accidental whitespace from headers to prevent KeyError exceptions
    df.columns = [c.strip() for c in df.columns]
    
    # 2. MERGE AUTHORS (Author 1 - Author 15)
    # Iterates through OJS author columns to create a single, clean citation string
    def join_authors(row):
        authors_list = []
        for i in range(1, 16):
            given_col = f"Given Name (Author {i})"
            family_col = f"Family Name (Author {i})"
            
            # Handle edge case: Single Author OJS export format
            if i == 1 and given_col not in row.index and "Given Name" in row.index:
                given_col = "Given Name"
                family_col = "Family Name"
            
            if given_col in row.index:
                given = str(row[given_col]).strip()
                family = str(row.get(family_col, '')).strip()
                
                # Filter out Pandas 'nan' artifacts
                if given.lower() == 'nan': given = ""
                if family.lower() == 'nan': family = ""
                
                full_name = f"{given} {family}".strip()
                if full_name:
                    authors_list.append(full_name)
        
        return ", ".join(authors_list) if authors_list else "Unknown Author"

    # Generate the unified authors column required by the Jinja2 template
    if 'Authors_Combined' not in df.columns:
        df['Authors_Combined'] = df.apply(join_authors, axis=1)

    # 3. HANDLE PAGE NUMBERS
    # Intelligently search for the page column regardless of OJS version naming
    page_col_found = next((c for c in df.columns if c.lower() in ['pages', 'range', 'page numbers']), None)
    
    if page_col_found:
        df['Pages_Clean'] = df[page_col_found].fillna("").astype(str).replace("nan", "")
    else:
        df['Pages_Clean'] = "" 

    # 4. HANDLE SECTION NAMES
    # Ensure every article is assigned a section for accurate ToC sorting
    section_col = next((c for c in df.columns if c.lower() in ['section', 'section title', 'discipline']), None)
    if section_col:
        df['Section Name'] = df[section_col].fillna("Articles")
    else:
        df['Section Name'] = "Articles"

    return df