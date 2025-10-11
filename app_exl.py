import os
import streamlit as st
import pandas as pd

# LangChain imports
from langchain.text_splitter import CharacterTextSplitter
from langchain.embeddings import OpenAIEmbeddings
from langchain.vectorstores import FAISS
from langchain.chains import RetrievalQA
from langchain.chat_models import ChatOpenAI

# -------------------------------
# Streamlit Setup
# -------------------------------
st.set_page_config(page_title="Agentic AI – Excel SOP Query", layout="wide")
st.title("Agentic AI – Excel SOP Query Interface")

# -------------------------------
# Paths
# -------------------------------
base_dir = os.path.dirname(__file__)
output_dir = os.path.join(base_dir, "data")
st.write("Using Excel folder:", output_dir)

# -------------------------------
# Load Excel files
# -------------------------------
excel_files = [f for f in os.listdir(output_dir) if f.endswith(".xlsx")]
if not excel_files:
    st.warning("No Excel files found in the 'output/' folder.")
else:
    st.write("Excel files found:", excel_files)

# -------------------------------
# Flatten Excel sheets into text per role
# -------------------------------
role_texts = {}
for file in excel_files:
    xls = pd.ExcelFile(os.path.join(output_dir, file))
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        # Flatten sheet into a single string
        role_texts[f"{file} - {sheet}"] = "\n".join(df.astype(str).values.flatten())

# -------------------------------
# User selects role
# -------------------------------
roles = list(role_texts.keys())
if roles:
    selected_role = st.selectbox("Select a role to query:", roles)

    query = st.text_input("Ask a question about SOP:")

    if query:
        # -------------------------------
        # Split text into chunks
        # -------------------------------
        text = role_texts[selected_role]
        splitter = CharacterTextSplitter(chunk_size=500, chunk_overlap=50)
        chunks = splitter.split_text(text)

        # -------------------------------
        # Create embeddings and retriever
        # -------------------------------
        embeddings = OpenAIEmbeddings()  # Make sure OPENAI_API_KEY is set
        vectorstore = FAISS.from_texts(chunks, embeddings)
        retriever = vectorstore.as_retriever()

        # -------------------------------
        # Connect to LLM via RetrievalQA
        # -------------------------------
        qa = RetrievalQA.from_chain_type(
            llm=ChatOpenAI(temperature=0.2),
            chain_type="stuff",
            retriever=retriever,
            return_source_documents=False
        )

        # -------------------------------
        # Get AI response
        # -------------------------------
        response = qa.run(query)
        st.markdown("**Agentic AI says:**")
        st.write(response)
