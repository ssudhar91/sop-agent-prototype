import os
import pickle
import streamlit as st
import pandas as pd

# LangChain & OpenAI imports
from langchain.text_splitter import CharacterTextSplitter
from langchain.embeddings import OpenAIEmbeddings
from langchain.vectorstores import FAISS
from langchain.chains import RetrievalQA
from langchain.chat_models import ChatOpenAI

st.set_page_config(page_title="Agentic AI – SOP Query", layout="wide")
st.title("Agentic AI – SOP Query Interface")

# -------------------------------
# 1️⃣ Set up KB paths
# -------------------------------
base_dir = os.path.dirname(__file__)
output_dir = os.path.join(base_dir, "output")
kb_pickle = os.path.join(base_dir, "preprocessed_kb.pkl")

# -------------------------------
# 2️⃣ Load or preprocess KB
# -------------------------------
if os.path.exists(kb_pickle):
    with open(kb_pickle, "rb") as f:
        kb_data = pickle.load(f)
    st.info("Loaded preprocessed KB successfully.")
else:
    st.info("Preprocessing KB from files...")
    kb_data = {}
    
    for file in os.listdir(output_dir):
        if file.endswith(".txt"):
            with open(os.path.join(output_dir, file), "r") as f:
                text = f.read()
            splitter = CharacterTextSplitter(chunk_size=500, chunk_overlap=50)
            kb_data[file] = splitter.split_text(text)
        elif file.endswith(".xlsx"):
            xls = pd.ExcelFile(os.path.join(output_dir, file))
            for sheet in xls.sheet_names:
                df = xls.parse(sheet)
                text = "\n".join(df.astype(str).values.flatten())
                splitter = CharacterTextSplitter(chunk_size=500, chunk_overlap=50)
                kb_data[f"{file}-{sheet}"] = splitter.split_text(text)
    
    with open(kb_pickle, "wb") as f:
        pickle.dump(kb_data, f)
    st.success("KB preprocessing completed and saved.")

# -------------------------------
# 3️⃣ Create or load embeddings
# -------------------------------
embeddings_pickle = os.path.join(base_dir, "vectorstores.pkl")

if os.path.exists(embeddings_pickle):
    with open(embeddings_pickle, "rb") as f:
        vectorstores = pickle.load(f)
    st.info("Loaded vectorstore embeddings successfully.")
else:
    st.info("Creating embeddings for KB...")
    embeddings = OpenAIEmbeddings()  # Ensure your API key is set in env
    vectorstores = {}
    for key, chunks in kb_data.items():
        vectorstores[key] = FAISS.from_texts(chunks, embeddings)
    with open(embeddings_pickle, "wb") as f:
        pickle.dump(vectorstores, f)
    st.success("Embeddings created and saved.")

# -------------------------------
# 4️⃣ User query interface
# -------------------------------
query = st.text_input("Ask something about any role or SOP:")

if query:
    # Merge all vectorstores for search (or select specific role)
    all_indices = list(vectorstores.values())
    
    # For demo, combine all indices into one retriever
    combined_docs = FAISS.merge_from([v for v in all_indices])
    retriever = combined_docs.as_retriever()

    qa = RetrievalQA.from_chain_type(
        llm=ChatOpenAI(temperature=0),
        chain_type="stuff",
        retriever=retriever
    )

    response = qa.run(query)
    st.markdown("**Agentic AI says:**")
    st.write(response)
