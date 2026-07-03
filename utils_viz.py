import pandas as pd
import plotly.express as px
from wordcloud import WordCloud
from sklearn.feature_extraction.text import CountVectorizer, ENGLISH_STOP_WORDS

def generate_wordcloud(df):
    """
    Generates a high-resolution Word Cloud from Article Titles using Bigrams.
    Aligns with FYP Objective 3.4.1: Co-Word Analysis (Keyword Clouds).
    Returns a raw RGB pixel array to bypass Streamlit's native image resizing bugs.
    """
    if 'Article Title' not in df.columns:
        return None
    
    titles = df['Article Title'].dropna().astype(str).tolist()
    
    if not titles:
        return None

    # Custom stop words to filter out generic academic terms
    custom_stopwords = list(ENGLISH_STOP_WORDS) + [
        'using', 'based', 'approach', 'method', 'analysis', 
        'system', 'model', 'study', 'evaluation', 'case', 'paper', 'proposed', 
        'editorial', 'preview'
    ]

    try:
        # Extract bigrams (2-word phrases) for better context 
        vectorizer = CountVectorizer(ngram_range=(2, 2), stop_words=custom_stopwords)
        X = vectorizer.fit_transform(titles)
        
        frequencies = dict(zip(vectorizer.get_feature_names_out(), X.toarray().sum(axis=0)))
        
        if not frequencies: 
            return None

        # Generate the Word Cloud at a massive 2K resolution for crisp rendering
        wordcloud = WordCloud(
            width=1600, 
            height=800, 
            background_color='white',
            colormap='viridis' 
        ).generate_from_frequencies(frequencies)
        
        # Return raw RGB pixel data (100% crash-proof in Streamlit)
        return wordcloud.to_array()
        
    except ValueError:
        return None

def plot_sections_pie(df):
    """
    Generates a premium Plotly Donut Chart showing the distribution of articles by section.
    """
    if 'Section Name' not in df.columns:
        return None
        
    counts = df['Section Name'].value_counts().reset_index()
    counts.columns = ['Section', 'Count']
    
    fig = px.pie(
        counts, 
        values='Count', 
        names='Section', 
        hole=0.45,
        color_discrete_sequence=px.colors.qualitative.Pastel 
    )
    
    fig.update_traces(
        textposition='outside', 
        textinfo='percent', 
        textfont_size=14,
        insidetextorientation='horizontal'
    )
    
    # Let Plotly handle the layout natively with a horizontal legend
    fig.update_layout(
        height=600, # Forces the chart canvas to be taller so the donut can expand
        title_text='<b>Issue Composition</b>',
        title_x=0.5, 
        title_y=0.95, 
        showlegend=True, 
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.2, 
            xanchor="center",
            x=0.5,
            font=dict(size=14)
        ),
        margin=dict(t=60, b=120, l=20, r=20),
        paper_bgcolor="rgba(0,0,0,0)", 
    )
    return fig

def plot_top_authors(df):
    """
    Generates a sleek Plotly Horizontal Bar Chart of the most frequent authors,
    automatically filtering out administrative/editorial entries.
    """
    if 'Authors' not in df.columns:
        return None
        
    # Filter out 'editorial preview' rows so editors aren't counted as authors
    if 'Article Title' in df.columns:
        mask = ~df['Article Title'].astype(str).str.contains('editorial preview', case=False, na=False)
        df_filtered = df[mask]
    else:
        df_filtered = df
        
    all_authors = df_filtered['Authors'].dropna().str.split(',').explode().str.strip()
    top_authors = all_authors.value_counts().head(10).reset_index()
    top_authors.columns = ['Author', 'Publications']
    
    # Sort ascending so the highest value appears at the top of the horizontal bar chart
    top_authors = top_authors.sort_values('Publications', ascending=True)
    
    fig = px.bar(
        top_authors, 
        x='Publications', 
        y='Author', 
        orientation='h', 
        color='Publications', 
        color_continuous_scale='Blues'
    )
    
    # Clean up the axes and hide the unnecessary color scale bar
    fig.update_layout(
        height=600, # Matches the height of the donut chart for UI symmetry
        font=dict(size=14), 
        title_text='<b>Top Contributors</b>',
        title_x=0.5,
        coloraxis_showscale=False, 
        xaxis_title="", 
        yaxis_title="",
        margin=dict(t=50, b=20, l=20, r=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=True, gridcolor='lightgray', tickfont=dict(size=14)),
        yaxis=dict(tickfont=dict(size=14))
    )
    return fig