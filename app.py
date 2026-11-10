"""Multi-Document Chat App - CDV3400 CA-3 (Generative AI on the Cloud).
 
Upload PDF, Word, PowerPoint, CSV, JSON, YAML or text files, index them in a FAISS
vector store using Gemini embeddings, and chat with them. Answers stream in from Gemini
and cite the file and the page / slide / rows used.
 
Heavy libraries are imported inside the functions that need them, so the page appears
immediately instead of waiting for every library to load at startup.
"""
 
from __future__ import annotations
 
import csv
import io
import json
import os
import shutil
import struct
import subprocess
import tempfile
import threading
import uuid
from pathlib import Path
 
import streamlit as st
from dotenv import dotenv_values
 
# Read GOOGLE_API_KEY and GEMINI_MODEL from the .env file. A filled-in value always wins
# over an old value left in the terminal; blank lines never erase a real value.
for _name, _value in dotenv_values(Path(__file__).parent / ".env").items():
    if _value:
        os.environ[_name] = _value
 
CHUNK_SIZE = 3000          # characters per chunk (neighbouring slides/pages are joined up to this size)
CHUNK_OVERLAP = 300        # characters shared between neighbouring chunks
TOP_K = 3                  # chunks retrieved per question
HISTORY_TURNS = 6          # previous messages included for follow-up questions
MAX_TABLE_ROWS = 1000      # rows read per CSV file
ROWS_PER_SECTION = 20      # CSV rows grouped into one section
EMBED_BATCH = 80           # chunks embedded per minute (free tier allows 100)
RATE_LIMIT_WAIT = 50       # seconds to wait if Google still says "too many requests"
EMBEDDING_MODEL = "models/gemini-embedding-001"
DEFAULT_MODEL = "gemini-2.5-flash"
CHAT_MODEL = os.getenv("GEMINI_MODEL") or DEFAULT_MODEL
DATABASE_URL = os.getenv("DATABASE_URL", "")
SUPPORTED_TYPES = ["pdf", "txt", "csv", "json", "yaml", "yml", "doc", "docx", "ppt", "pptx"]
 
 
# ------------------------------------------------------------ file readers
# Every reader returns a list of (location, text) pairs.
def read_pdf(file):
    from pypdf import PdfReader
 
    return [
        (f"page {number}", page.extract_text() or "")
        for number, page in enumerate(PdfReader(file).pages, start=1)
    ]
 
 
def read_text(file):
    return [("text", file.read().decode("utf-8", errors="ignore"))]
 
 
def read_csv(file):
    text = file.read().decode("utf-8", errors="ignore")
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return []
    header, body = rows[0], rows[1 : MAX_TABLE_ROWS + 1]
    sections = []
    for start in range(0, len(body), ROWS_PER_SECTION):
        group = body[start : start + ROWS_PER_SECTION]
        lines = ["; ".join(f"{h}: {v}" for h, v in zip(header, row) if v) for row in group]
        sections.append((f"rows {start + 1}-{start + len(group)}", "\n".join(lines)))
    return sections
 
 
def flatten(data, prefix=""):
    """Turn nested JSON/YAML data into 'path: value' lines."""
    if isinstance(data, dict):
        return [line for key, value in data.items() for line in flatten(value, f"{prefix}{key}.")]
    if isinstance(data, list):
        return [line for i, value in enumerate(data) for line in flatten(value, f"{prefix}{i}.")]
    return [f"{prefix.rstrip('.')}: {data}"]
 
 
def read_json(file):
    return [("content", "\n".join(flatten(json.load(file))))]
 
 
def read_yaml(file):
    import yaml
 
    lines = [line for doc in yaml.safe_load_all(file) for line in flatten(doc)]
    return [("content", "\n".join(lines))]
 
 
def read_docx(file):
    from docx import Document as WordDocument
 
    document = WordDocument(file)
    paragraphs = [p.text for p in document.paragraphs]
    tables = [
        " | ".join(cell.text for cell in row.cells)
        for table in document.tables
        for row in table.rows
    ]
    return [("document", "\n".join(paragraphs + tables))]
 
 
def read_pptx(file):
    from pptx import Presentation
 
    sections = []
    for number, slide in enumerate(Presentation(file).slides, start=1):
        texts = [shape.text_frame.text for shape in slide.shapes if shape.has_text_frame]
        sections.append((f"slide {number}", "\n".join(texts)))
    return sections
 
 
def read_doc(file):
    """Read old binary .doc files with the catdoc tool."""
    if not shutil.which("catdoc"):
        raise RuntimeError("'catdoc' is not installed; save the file as .docx instead")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as temp:
        temp.write(file.read())
    try:
        result = subprocess.run(["catdoc", temp.name], capture_output=True, text=True, check=True)
    finally:
        os.remove(temp.name)
    return [("content", result.stdout)]
 
 
# Record types inside the binary PowerPoint (.ppt) format.
PPT_SLIDE, PPT_TEXT_CHARS, PPT_TEXT_BYTES = 0x03EE, 0x0FA0, 0x0FA8
 
 
def collect_ppt_text(data, start, end, texts):
    """Walk the records between start and end, collecting every text record."""
    position = start
    while position + 8 <= end:
        info, record_type, length = struct.unpack_from("<HHI", data, position)
        body = position + 8
        if info & 0x0F == 0x0F:  # a container: it holds more records
            collect_ppt_text(data, body, min(body + length, end), texts)
        elif record_type == PPT_TEXT_CHARS:
            texts.append(data[body : body + length].decode("utf-16-le", errors="ignore"))
        elif record_type == PPT_TEXT_BYTES:
            texts.append(data[body : body + length].decode("latin-1"))
        position = body + length
 
 
def read_ppt(file):
    """Read old binary .ppt files, one section per slide."""
    import olefile
 
    data = olefile.OleFileIO(file).openstream("PowerPoint Document").read()
    slides, position = [], 0
    while position + 8 <= len(data):
        _, record_type, length = struct.unpack_from("<HHI", data, position)
        if record_type == PPT_SLIDE:
            texts = []
            collect_ppt_text(data, position + 8, position + 8 + length, texts)
            slides.append(texts)
        position += 8 + length
    return [
        (f"slide {number}", "\n".join(texts).replace("\r", "\n"))
        for number, texts in enumerate(slides, start=1)
    ]
 
 
READERS = {
    ".pdf": read_pdf,
    ".txt": read_text,
    ".csv": read_csv,
    ".json": read_json,
    ".yaml": read_yaml,
    ".yml": read_yaml,
    ".docx": read_docx,
    ".pptx": read_pptx,
    ".doc": read_doc,
    ".ppt": read_ppt,
}
 
 
# ---------------------------------------------------------------- pipeline
def load_files(uploaded_files):
    """Read every uploaded file; return (documents, skipped-file messages)."""
    from langchain_core.documents import Document
 
    documents, skipped = [], []
    for upload in uploaded_files:
        upload.seek(0)
        reader = READERS[Path(upload.name).suffix.lower()]
        try:
            sections = reader(upload)
        except Exception as error:  # one bad file must not stop the others
            skipped.append(f"{upload.name}: {error}")
            continue
        readable = [(loc, text) for loc, text in sections if text.strip()]
        if not readable:
            skipped.append(f"{upload.name}: no readable text")
        for location, text in readable:
            documents.append(
                Document(
                    page_content=text,
                    metadata={"source": upload.name, "location": location},
                )
            )
    return documents, skipped
 
 
def merge_sections(documents):
    """Join neighbouring short sections of one file so there are fewer chunks to embed."""
    from langchain_core.documents import Document
 
    groups = []
    for doc in documents:
        source, location = doc.metadata["source"], doc.metadata["location"]
        group = groups[-1] if groups else None
        if (
            group
            and group["source"] == source
            and len(group["text"]) + len(doc.page_content) <= CHUNK_SIZE
        ):
            group["text"] += f"\n\n[{location}]\n{doc.page_content}"
            group["end"] = location
        else:
            groups.append(
                {
                    "source": source,
                    "start": location,
                    "end": location,
                    "text": f"[{location}]\n{doc.page_content}",
                }
            )
    return [
        Document(
            page_content=g["text"],
            metadata={
                "source": g["source"],
                "location": g["start"] if g["start"] == g["end"] else f"{g['start']} to {g['end']}",
            },
        )
        for g in groups
    ]
 
 
def split_documents(documents):
    """Merge short sections, then split long ones into overlapping chunks."""
    from langchain_text_splitters import RecursiveCharacterTextSplitter
 
    splitter = RecursiveCharacterTextSplitter(
        chunk_size=CHUNK_SIZE, chunk_overlap=CHUNK_OVERLAP
    )
    return splitter.split_documents(merge_sections(documents))
 
 
def build_vector_store(chunks, embeddings, progress=None):
    """Embed in batches and wait when the free-tier per-minute limit is hit."""
    import time
 
    from langchain_community.vectorstores import FAISS
 
    store = None
    for start in range(0, len(chunks), EMBED_BATCH):
        batch = chunks[start : start + EMBED_BATCH]
        began = time.time()
        while True:
            try:
                if store is None:
                    store = FAISS.from_documents(batch, embeddings)
                else:
                    store.add_documents(batch)
                break
            except Exception as error:
                if "429" not in str(error) and "RESOURCE_EXHAUSTED" not in str(error):
                    raise
                time.sleep(RATE_LIMIT_WAIT)
        if progress:
            progress(min(start + EMBED_BATCH, len(chunks)), len(chunks))
        if start + EMBED_BATCH < len(chunks):  # stay under 100 requests per minute
            time.sleep(max(0, 60 - (time.time() - began)))
    return store
 
 
def format_history(messages):
    """Turn the last few chat messages into plain text for the prompt."""
    recent = messages[-HISTORY_TURNS:]
    return "\n".join(f"{m['role'].capitalize()}: {m['content']}" for m in recent)
 
 
def initialize_database() -> bool:
    """Create the PostgreSQL tables when DATABASE_URL is configured."""
    if not DATABASE_URL:
        return False

    import psycopg

    with psycopg.connect(DATABASE_URL) as connection:
        connection.execute("""
            CREATE TABLE IF NOT EXISTS app_sessions (
                session_id TEXT PRIMARY KEY,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
        """)
        connection.execute("""
            CREATE TABLE IF NOT EXISTS documents (
                session_id TEXT NOT NULL REFERENCES app_sessions(session_id) ON DELETE CASCADE,
                file_name TEXT NOT NULL,
                section_count INTEGER NOT NULL DEFAULT 0,
                uploaded_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                PRIMARY KEY (session_id, file_name)
            )
        """)
        connection.execute("""
            CREATE TABLE IF NOT EXISTS chat_messages (
                id BIGSERIAL PRIMARY KEY,
                session_id TEXT NOT NULL REFERENCES app_sessions(session_id) ON DELETE CASCADE,
                role TEXT NOT NULL CHECK (role IN ('user', 'assistant')),
                content TEXT NOT NULL,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
        """)
    return True


def save_document_metadata(session_id: str, documents) -> None:
    """Persist uploaded document names and section counts for this app session."""
    import psycopg

    counts = {}
    for document in documents:
        name = document.metadata["source"]
        counts[name] = counts.get(name, 0) + 1

    with psycopg.connect(DATABASE_URL) as connection:
        connection.execute(
            "INSERT INTO app_sessions (session_id) VALUES (%s) ON CONFLICT DO NOTHING",
            (session_id,),
        )
        for name, count in counts.items():
            connection.execute(
                """
                INSERT INTO documents (session_id, file_name, section_count)
                VALUES (%s, %s, %s)
                ON CONFLICT (session_id, file_name)
                DO UPDATE SET section_count = EXCLUDED.section_count, uploaded_at = NOW()
                """,
                (session_id, name, count),
            )


def load_chat_history(session_id: str) -> list[dict]:
    """Restore saved messages after the user reprocesses their documents."""
    import psycopg

    with psycopg.connect(DATABASE_URL) as connection:
        connection.execute(
            "INSERT INTO app_sessions (session_id) VALUES (%s) ON CONFLICT DO NOTHING",
            (session_id,),
        )
        rows = connection.execute(
            "SELECT role, content FROM chat_messages WHERE session_id = %s ORDER BY id",
            (session_id,),
        ).fetchall()
    return [{"role": row[0], "content": row[1]} for row in rows]


def save_chat_exchange(session_id: str, question: str, answer: str) -> None:
    """Save one user question and its assistant response."""
    import psycopg

    with psycopg.connect(DATABASE_URL) as connection:
        connection.execute(
            "INSERT INTO app_sessions (session_id) VALUES (%s) ON CONFLICT DO NOTHING",
            (session_id,),
        )
        connection.executemany(
            "INSERT INTO chat_messages (session_id, role, content) VALUES (%s, %s, %s)",
            [
                (session_id, "user", question),
                (session_id, "assistant", answer),
            ],
        )


def build_prompt(question, sources, history):
    """Combine retrieved context, chat history and the question into one prompt."""
    context = "\n\n".join(
        f"[{d.metadata['source']}, {d.metadata['location']}]\n{d.page_content}"
        for d in sources
    )
    return (
        "You are a helpful assistant. Answer the question concisely using ONLY the "
        "context from the uploaded files below. If the answer is not in the context, "
        'say "The answer is not available in the uploaded documents."\n\n'
        f"Conversation so far:\n{history or '(none)'}\n\n"
        f"Context:\n{context}\n\n"
        f"Question: {question}\nAnswer:"
    )
 
 
def message_text(message):
    """Extract plain text from a model reply (content may be text or blocks)."""
    content = message.content
    if isinstance(content, str):
        return content
    return "".join(
        block.get("text", "") if isinstance(block, dict) else str(block)
        for block in content
    )
 
 
def stream_answer(question, store, llm, messages):
    """Retrieve relevant chunks; return (stream of answer text, sources)."""
    sources = store.similarity_search(question, k=TOP_K)
    prompt = build_prompt(question, sources, format_history(messages))
    return (message_text(part) for part in llm.stream(prompt)), sources
 
 
# ------------------------------------------------- cached models and warm-up
@st.cache_resource(show_spinner=False)
def get_llm(model: str):
    """Create the chat model once and reuse it; ask it to think less for speed."""
    from langchain_google_genai import ChatGoogleGenerativeAI
 
    settings = {"model": model, "temperature": 0.2}
    if model.startswith("gemini-2.5-flash"):
        settings["thinking_budget"] = 0        # no hidden thinking step
    elif model.startswith("gemini-3"):
        settings["thinking_level"] = "low"     # least thinking Gemini 3 models allow
    return ChatGoogleGenerativeAI(**settings)
 
 
@st.cache_resource(show_spinner=False)
def get_embeddings():
    """Create the embedding model once and reuse it."""
    from langchain_google_genai import GoogleGenerativeAIEmbeddings
 
    return GoogleGenerativeAIEmbeddings(model=EMBEDDING_MODEL)
 
 
@st.cache_resource(show_spinner=False)
def warm_up() -> bool:
    """Load the heavy libraries in the background so the first upload is fast."""
 
    def preload():
        try:
            import faiss  # noqa: F401
            import langchain_community.vectorstores  # noqa: F401
            import langchain_core.documents  # noqa: F401
            import langchain_google_genai  # noqa: F401
            import langchain_text_splitters  # noqa: F401
            import pypdf  # noqa: F401
        except Exception:
            pass  # warm-up is only an optimisation
 
    threading.Thread(target=preload, daemon=True).start()
    return True
 
 
# ---------------------------------------------------------------------- UI
def process_documents(uploaded_files) -> None:
    """Read, split and index the uploaded files; keep the result in the session."""
    documents, skipped = load_files(uploaded_files)
    for message in skipped:
        st.sidebar.warning(f"Skipped {message}")
    if not documents:
        st.sidebar.error("No readable text found in the uploaded files.")
        return
    chunks = split_documents(documents)
    bar = st.sidebar.progress(0.0, text="Embedding your documents...")
    st.session_state.store = build_vector_store(
        chunks,
        get_embeddings(),
        lambda done, total: bar.progress(done / total, text=f"Embedded {done} of {total} chunks"),
    )
    bar.empty()
    st.session_state.stats = {
        "files": len({d.metadata["source"] for d in documents}),
        "sections": len(documents),
        "chunks": len(chunks),
    }
    if DATABASE_URL:
        try:
            save_document_metadata(st.session_state.session_id, documents)
            st.sidebar.caption("Document details saved to PostgreSQL.")
            st.session_state.messages = load_chat_history(st.session_state.session_id)
        except Exception as error:
            st.sidebar.warning(f"Database save failed: {error}")
    else:
        st.session_state.messages = []
 
 
def render_sidebar() -> None:
    """Sidebar: upload files, process them, and show index statistics."""
    st.sidebar.header("📚 Your documents")
    uploaded_files = st.sidebar.file_uploader(
        "Upload one or more files",
        type=SUPPORTED_TYPES,
        accept_multiple_files=True,
    )
    if st.sidebar.button("Process documents", type="primary", disabled=not uploaded_files):
        with st.spinner("Reading and indexing your files..."):
            try:
                process_documents(uploaded_files)
            except Exception as error:  # surface API / network problems to the user
                st.sidebar.error(f"Processing failed: {error}")
 
    stats = st.session_state.get("stats")
    if stats:
        st.sidebar.success("Documents are ready. Ask your questions!")
        st.sidebar.metric("Files", stats["files"])
        st.sidebar.metric("Sections (pages / slides / rows)", stats["sections"])
 
    if st.sidebar.button("Clear chat"):
        st.session_state.messages = []
 
 
def render_chat() -> None:
    """Main area: chat history, input box and answers with their sources."""
    for message in st.session_state.messages:
        st.chat_message(message["role"]).write(message["content"])
 
    question = st.chat_input("Ask a question about your files")
    if not question:
        return
 
    st.chat_message("user").write(question)
    with st.chat_message("assistant"):
        try:
            llm = get_llm(CHAT_MODEL)
            with st.spinner("Searching your documents..."):
                stream, sources = stream_answer(
                    question, st.session_state.store, llm, st.session_state.messages
                )
            answer = st.write_stream(stream)
        except Exception as error:
            st.error(f"Could not get an answer: {error}")
            return
        with st.expander("Sources used"):
            for doc in sources:
                st.markdown(
                    f"**{doc.metadata['source']}, {doc.metadata['location']}**\n\n"
                    f"{doc.page_content[:400]}..."
                )
 
    st.session_state.messages += [
        {"role": "user", "content": question},
        {"role": "assistant", "content": answer},
    ]
    if DATABASE_URL:
        try:
            save_chat_exchange(st.session_state.session_id, question, answer)
        except Exception as error:
            st.warning(f"Could not save chat history to PostgreSQL: {error}")
 
 
def apply_theme() -> None:
    """Apply a clean orange-and-white theme to Streamlit's interface."""
    st.markdown(
        """
        <style>
        :root {
            --orange: #f97316;
            --orange-dark: #ea580c;
            --orange-soft: #fff7ed;
            --ink: #242424;
            --muted: #777777;
            --line: #f0e5dc;
        }
        .stApp {
            background: #fffdfb;
            color: var(--ink);
        }
        [data-testid="stHeader"] {
            background: rgba(255, 253, 251, 0.96);
        }
        [data-testid="stDeployButton"] { display: none !important; }
        [data-testid="stToolbar"] { right: 1rem; }
        [data-testid="stSidebar"] {
            background: #fff7ed;
            border-right: 1px solid #f3e5d8;
        }
        [data-testid="stSidebar"] > div:first-child {
            padding-top: 2rem;
        }
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: #292524;
            font-weight: 700;
        }
        [data-testid="stFileUploader"] section {
            background: #ffffff;
            border: 1px dashed #f4a261;
            border-radius: 14px;
        }
        [data-testid="stFileUploader"] section:hover {
            border-color: var(--orange);
        }
        .stButton > button[kind="primary"],
        [data-testid="stFormSubmitButton"] > button {
            background: var(--orange);
            color: #ffffff;
            border: 1px solid var(--orange);
            border-radius: 10px;
            font-weight: 650;
            min-height: 2.7rem;
        }
        .stButton > button[kind="primary"]:hover,
        [data-testid="stFormSubmitButton"] > button:hover {
            background: var(--orange-dark);
            border-color: var(--orange-dark);
            color: #ffffff;
        }
        .stButton > button:not([kind="primary"]) {
            border: 1px solid #ead8ca;
            border-radius: 10px;
            color: #44403c;
            background: #ffffff;
        }
        [data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #f1e4d9;
            border-radius: 12px;
            padding: 12px 14px;
        }
        [data-testid="stMetricValue"] {
            color: var(--orange-dark);
            font-weight: 750;
        }
        [data-testid="stChatMessage"] {
            background: #ffffff;
            border: 1px solid #f0e7df;
            border-radius: 16px;
            padding: 1rem;
            margin-bottom: 0.8rem;
        }
        [data-testid="stChatInput"] {
            border-color: #edc6a8;
            border-radius: 14px;
        }
        [data-testid="stChatInput"]:focus-within {
            border-color: var(--orange);
            box-shadow: 0 0 0 2px #ffedd5;
        }
        .app-hero {
            padding: 1.2rem 0 1.5rem 0;
            border-bottom: 1px solid #f2e8e0;
            margin-bottom: 1.8rem;
        }
        .app-kicker {
            display: inline-block;
            color: #c2410c;
            background: #ffedd5;
            border-radius: 999px;
            padding: 5px 11px;
            font-size: 0.78rem;
            font-weight: 700;
            letter-spacing: .02em;
            margin-bottom: 0.7rem;
        }
        .app-title {
            color: #292524;
            font-size: clamp(1.8rem, 3vw, 2.5rem);
            font-weight: 800;
            letter-spacing: -0.04em;
            line-height: 1.15;
            margin: 0;
        }
        .app-subtitle {
            color: #78716c;
            font-size: 1rem;
            margin-top: 0.6rem;
        }
        .welcome-card {
            background: linear-gradient(135deg, #fff7ed 0%, #ffffff 75%);
            border: 1px solid #f3dfce;
            border-radius: 20px;
            padding: 2rem;
            margin: 1rem 0 1.5rem 0;
        }
        .welcome-card h3 {
            color: #292524;
            font-size: 1.35rem;
            margin: 0 0 .5rem 0;
        }
        .welcome-card p {
            color: #78716c;
            margin: 0;
            line-height: 1.6;
        }
        @media (max-width: 700px) {
            [data-testid="stSidebar"] { min-width: 260px; }
            .welcome-card { padding: 1.25rem; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(
        page_title="Chat with Your Documents",
        page_icon="📚",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_theme()
    st.markdown(
        """
        <div class="app-hero">
            <div class="app-kicker">YOUR PERSONAL DOCUMENT ASSISTANT</div>
            <h1 class="app-title">📚 Chat with your documents</h1>
            <p class="app-subtitle">Upload your files, ask questions, and get answers grounded in your content.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not os.getenv("GOOGLE_API_KEY"):
        st.error("GOOGLE_API_KEY is missing. Add it to the .env file next to app.py, then restart.")
        st.stop()

    warm_up()
    st.session_state.setdefault("messages", [])
    # Keep a stable session key in the URL so saved chat history can be reloaded
    # after a browser refresh. This is a lightweight single-user/session approach.
    query_session_id = st.query_params.get("session_id")
    if query_session_id:
        st.session_state.session_id = query_session_id
    else:
        st.session_state.setdefault("session_id", str(uuid.uuid4()))
        st.query_params["session_id"] = st.session_state.session_id
    if DATABASE_URL:
        try:
            initialize_database()
        except Exception as error:
            st.sidebar.warning(f"PostgreSQL is not available: {error}")
    render_sidebar()

    if "store" not in st.session_state:
        st.markdown(
            """
            <div class="welcome-card">
                <h3>Start with your documents</h3>
                <p>Upload one or more PDF, Word, PowerPoint, CSV, JSON, YAML, or text files
                from the left panel. Select <b>Process documents</b> and your assistant
                will be ready to answer questions.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.stop()

    if not st.session_state.messages:
        st.markdown(
            """
            <div class="welcome-card">
                <h3>Your documents are ready. What would you like to know?</h3>
                <p>Ask for a summary, find a specific concept, compare topics, or get an
                explanation based only on the files you uploaded.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    render_chat()
 
 
if __name__ == "__main__":
    main()
 