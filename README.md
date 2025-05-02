import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import streamlit as st

# --- Function to add slide title box ---
def add_slide_title(slide, title_text):
    left = Inches(0.5)
    top = Inches(0.2)
    width = Inches(9)
    height = Inches(0.6)

    title_box = slide.shapes.add_textbox(left, top, width, height)
    text_frame = title_box.text_frame
    text_frame.text = title_text
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Styling
    run = text_frame.paragraphs[0].runs[0]
    run.font.bold = True
    run.font.size = Pt(20)

# --- Function to generate graph ---
def generate_graph(df, x_col, y_col, graph_type, title):
    if graph_type == 'Bar':
        fig = px.bar(df, x=x_col, y=y_col, title=title)
    elif graph_type == 'Pie':
        fig = px.pie(df, names=x_col, values=y_col, title=title)
    else:
        st.error("Please select a valid graph type (Bar or Pie).")
        return None
    return fig

# --- Function to save graph to PPT ---
def save_graph_to_ppt(graph, ppt, slide_layout, title, graph_path):
    fig = graph
    fig.write_image(graph_path)

    # Create a new slide and add title
    slide = ppt.slides.add_slide(slide_layout)
    add_slide_title(slide, title)
    left = Inches(0.5)
    top_image = Inches(0.9)
    height = Inches(6.5)
    slide.shapes.add_picture(graph_path, left, top_image, height=height)

# --- Streamlit UI Setup ---
st.title("Excel Data to PPT Graph Generator")

# Upload Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write(df.head())  # Display a sample of the data

    # Let user choose x and y axis columns
    x_col = st.selectbox("Select X Axis Column", df.columns)
    y_col = st.selectbox("Select Y Axis Column", df.columns)
    graph_type = st.selectbox("Select Graph Type", ["Bar", "Pie"])

    if st.button("Generate PPT"):
        # Initialize PPT
        ppt = Presentation()
        slide_layout = ppt.slide_layouts[5]  # blank slide layout

        # Generate graph
        title = f'{y_col} vs {x_col} ({graph_type} Chart)'
        graph = generate_graph(df, x_col, y_col, graph_type, title)
        if graph:
            graph_path = 'temp_graph.png'
            save_graph_to_ppt(graph, ppt, slide_layout, title, graph_path)
            ppt.save('output_presentation.pptx')
            st.success("PPT generated successfully!")
            st.download_button("Download PPT", "output_presentation.pptx")
