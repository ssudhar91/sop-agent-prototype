import os
import pickle
import streamlit as st
import pandas as pd

# LangChain & OpenAI
from langchain.text_splitter import CharacterTextSplitter
from langchain.embeddings import OpenAIEmbeddings
from langchain.vectorstores import FAISS
from langchain.chains import ConversationalRetrievalChain
from langchain.chat_models import ChatOpenAI

# -------------------------------
# 1️⃣ Streamlit Setup
# -------------------------------
st.set_page_config(page_title="Agentic AI – SOP Query", layout="wide")
st.title("Agentic AI – SOP Query Interface")

# -------------------------------
# 2️⃣ Setup paths (Codespace-safe)
# -------------------------------
base_dir = os.path.dirname(__file__)
output_dir = os.path.join(base_dir, "output")
kb_pickle = os.path.join(base_dir, "preprocessed_kb.pkl")
embeddings_pickle = os.path.join(base_dir, "vectorstores.pkl")

st.write("Output folder being used:", output_dir)

# -------------------------------
# 3️⃣ Load or preprocess KB
# -------------------------------
if os.path.exists(kb_pickle):
    with open(kb_pickle, "rb") as f:
        kb_data = pickle.load(f)
    st.info("Loaded preprocessed KB successfully.")
else:
    st.info("Preprocessing KB from files...")
    kb_data = {}
    
    for file in os.listdir(output_dir):
        file_path = os.path.join(output_dir, file)
        if file.endswith(".txt"):
            with open(file_path, "r") as f:
                text = f.read()
            splitter = CharacterTextSplitter(chunk_size=500, chunk_overlap=50)
            kb_data[file] = splitter.split_text(text)
        elif file.endswith(".xlsx"):
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                df = xls.parse(sheet)
                text = "\n".join(df.astype(str).values.flatten())
                splitter = CharacterTextSplitter(chunk_size=500, chunk_overlap=50)
                kb_data[f"{file}-{sheet}"] = splitter.split_text(text)
    
    with open(kb_pickle, "wb") as f:
        pickle.dump(kb_data, f)
    st.success("KB preprocessing completed and saved.")

# -------------------------------
# 4️⃣ Create or load embeddings
# -------------------------------
if os.path.exists(embeddings_pickle):
    with open(embeddings_pickle, "rb") as f:
        vectorstores = pickle.load(f)
    st.info("Loaded vectorstore embeddings successfully.")
else:
    st.info("Creating embeddings for KB...")
    embeddings = OpenAIEmbeddings()  # Make sure OPENAI_API_KEY is set
    vectorstores = {}
    for key, chunks in kb_data.items():
        vectorstores[key] = FAISS.from_texts(chunks, embeddings)
    with open(embeddings_pickle, "wb") as f:
        pickle.dump(vectorstores, f)
    st.success("Embeddings created and saved.")

# -------------------------------
# 5️⃣ User selects a role (optional) and queries SOPs
# -------------------------------
roles = list(vectorstores.keys())
selected_role = st.selectbox("Select a role (optional, 'All' searches all SOPs):", ["All"] + roles)
query = st.text_input("Ask a question about SOPs:")

if query:
    # Choose retriever
    if selected_role == "All":
        all_indices = list(vectorstores.values())
        combined = FAISS.merge_from(all_indices)
        retriever = combined.as_retriever()
    else:
        retriever = vectorstores[selected_role].as_retriever()
    
    # -------------------------------
    # 6️⃣ Conversational Retrieval Chain
    # -------------------------------
    qa_chain = ConversationalRetrievalChain.from_llm(
        llm=ChatOpenAI(temperature=0.2),  # factual but natural
        retriever=retriever,
        return_source_documents=False
    )

    response = qa_chain.run(query)
    st.markdown("**Agentic AI says:**")
    st.write(response)

    # Optional: debug retrieved chunks
    # docs = retriever.get_relevant_documents(query)
    # st.write("Top retrieved chunks (debug):")
    # for doc in docs[:3]:
    #     st.write(doc.page_content)
