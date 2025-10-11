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
# Path to master Excel file
# -------------------------------
base_dir = os.path.dirname(__file__)
excel_file_path = os.path.join(base_dir, "data", "Novotech_SOP_Matrix.xlsx")

if not os.path.exists(excel_file_path):
    st.error(f"Master Excel file not found at {excel_file_path}")
else:
    st.write(f"Using master Excel file: {excel_file_path}")

    # -------------------------------
    # Load Excel file and flatten sheets into text per role
    # -------------------------------
    xls = pd.ExcelFile(excel_file_path)
    role_texts = {}
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        role_texts[sheet] = "\n".join(df.astype(str).values.flatten())

    # -------------------------------
    # User selects role
    # -------------------------------
    roles = list(role_texts.keys())
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
