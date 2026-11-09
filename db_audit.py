import os
import json
import sqlite3
import difflib
from datetime import datetime
from typing import List, Dict, Optional, Any
import boto3
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

DB_FILE = os.path.join(os.path.dirname(__file__), "recon_agent.db")
EMBEDDING_MODEL_ID = os.getenv("BEDROCK_EMBEDDING_MODEL_ID", "amazon.titan-embed-text-v2:0")

# Initialize Bedrock client for embeddings
def get_bedrock_client():
    return boto3.client(
        "bedrock-runtime",
        region_name=os.getenv("AWS_REGION", "us-east-1"),
        aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
        aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
    )


def get_db_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Initializes SQLite schema for audit logging and RAG memory."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id TEXT,
                txn_id TEXT NOT NULL,
                txn_date TEXT,
                memo TEXT NOT NULL,
                amount REAL,
                ai_vendor TEXT,
                ai_gl_account TEXT,
                ai_gl_description TEXT,
                final_vendor TEXT,
                final_gl_account TEXT,
                final_gl_description TEXT,
                confidence REAL,
                status TEXT NOT NULL,
                reasoning TEXT,
                model_id TEXT,
                source TEXT DEFAULT 'AI',
                rag_sources_used TEXT,
                embedding TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_txn_id ON audit_log(txn_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_status ON audit_log(status)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_memo ON audit_log(memo)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_batch ON audit_log(batch_id)")
        conn.commit()


def get_embedding(text: str) -> Optional[List[float]]:
    """Generates 256-dimensional normalized vector embedding using Amazon Titan Embeddings v2."""
    if not text or not text.strip():
        return None
    try:
        client = get_bedrock_client()
        body = json.dumps({
            "inputText": text.strip()[:1000],
            "dimensions": 256,
            "normalize": True
        })
        response = client.invoke_model(
            modelId=EMBEDDING_MODEL_ID,
            body=body,
            contentType="application/json",
            accept="application/json"
        )
        response_body = json.loads(response["body"].read())
        return response_body.get("embedding")
    except Exception as e:
        print(f"[Embedding Warning] Bedrock embedding call failed: {e}")
        return None


def cosine_similarity(v1: List[float], v2: List[float]) -> float:
    """Computes cosine similarity between two unit/normalized vectors."""
    if not v1 or not v2 or len(v1) != len(v2):
        return 0.0
    return sum(a * b for a, b in zip(v1, v2))


def log_reconciliation_decision(
    batch_id: str,
    txn_id: str,
    txn_date: str,
    memo: str,
    amount: float,
    ai_vendor: Optional[str],
    ai_gl_account: Optional[str],
    ai_gl_description: Optional[str],
    confidence: float,
    status: str,
    reasoning: str,
    model_id: str,
    rag_sources_used: Optional[List[Dict[str, Any]]] = None,
    source: str = "AI",
    generate_embed: bool = True
) -> int:
    """Inserts a traceable reconciliation decision into the SQLite audit log."""
    embedding = None
    if generate_embed:
        embedding = get_embedding(memo)
    
    embedding_json = json.dumps(embedding) if embedding else None
    rag_json = json.dumps(rag_sources_used) if rag_sources_used else None
    now_str = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")

    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO audit_log (
                batch_id, txn_id, txn_date, memo, amount,
                ai_vendor, ai_gl_account, ai_gl_description,
                final_vendor, final_gl_account, final_gl_description,
                confidence, status, reasoning, model_id,
                source, rag_sources_used, embedding,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            batch_id, txn_id, txn_date, memo, amount,
            ai_vendor, ai_gl_account, ai_gl_description,
            ai_vendor, ai_gl_account, ai_gl_description,
            confidence, status, reasoning, model_id,
            source, rag_json, embedding_json,
            now_str, now_str
        ))
        conn.commit()
        return cursor.lastrowid


def record_human_review(
    txn_id: str,
    final_vendor: str,
    final_gl_account: str,
    final_gl_description: str,
    new_status: str,
    reasoning: Optional[str] = None,
    batch_id: Optional[str] = None
):
    """
    Records a human accountant review or override.
    Updates the audit record, marks source as 'HUMAN', and flags it as a priority precedent for future RAG memory.
    """
    now_str = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    with get_db_connection() as conn:
        cursor = conn.cursor()
        # Find latest record for this txn_id (or batch)
        query = "SELECT id, memo, embedding FROM audit_log WHERE txn_id = ?"
        params = [txn_id]
        if batch_id:
            query += " AND batch_id = ?"
            params.append(batch_id)
        query += " ORDER BY id DESC LIMIT 1"

        cursor.execute(query, params)
        row = cursor.fetchone()

        if row:
            rec_id = row["id"]
            memo = row["memo"]
            embedding_json = row["embedding"]

            # If embedding was missing, compute now
            if not embedding_json and memo:
                emb = get_embedding(memo)
                if emb:
                    embedding_json = json.dumps(emb)

            override_reason = reasoning or "Human accountant approved/overrode AI categorization"
            cursor.execute("""
                UPDATE audit_log
                SET final_vendor = ?,
                    final_gl_account = ?,
                    final_gl_description = ?,
                    status = ?,
                    reasoning = ?,
                    source = 'HUMAN',
                    embedding = COALESCE(embedding, ?),
                    updated_at = ?
                WHERE id = ?
            """, (
                final_vendor, final_gl_account, final_gl_description,
                new_status, override_reason, embedding_json, now_str, rec_id
            ))
            conn.commit()


def retrieve_rag_memory(
    memo: str,
    top_k: int = 3,
    min_similarity: float = 0.45,
    exclude_txn_id: Optional[str] = None
) -> List[Dict[str, Any]]:
    """
    RAG Memory Retriever:
    Retrieves the most relevant past reconciled decisions from the SQLite audit database.
    Prioritizes approved / human-overridden decisions.
    Combines Bedrock semantic vector similarity + token/lexical fuzzy similarity.
    """
    if not memo or not memo.strip():
        return []

    memo_clean = memo.strip()
    query_emb = get_embedding(memo_clean)

    with get_db_connection() as conn:
        cursor = conn.cursor()
        # Fetch precedent records that have been approved or reviewed
        # Priority statuses: 'AUTO_APPROVE', 'HUMAN_APPROVED', 'HUMAN_OVERRIDE'
        cursor.execute("""
            SELECT id, txn_id, txn_date, memo, amount,
                   COALESCE(final_vendor, ai_vendor) AS vendor_name,
                   COALESCE(final_gl_account, ai_gl_account) AS gl_account,
                   COALESCE(final_gl_description, ai_gl_description) AS gl_description,
                   confidence, status, reasoning, source, embedding, updated_at
            FROM audit_log
            WHERE vendor_name IS NOT NULL
              AND gl_account IS NOT NULL
              AND status IN ('AUTO_APPROVE', 'HUMAN_APPROVED', 'HUMAN_OVERRIDE', 'REVIEW')
            ORDER BY id DESC
            LIMIT 200
        """)
        rows = cursor.fetchall()

    if not rows:
        return []

    scored_candidates = []
    seen_memos = set()

    for row in rows:
        past_id = row["txn_id"]
        if exclude_txn_id and past_id == exclude_txn_id:
            continue

        past_memo = row["memo"]
        # Deduplicate identical past memo precedents
        norm_key = f"{past_memo.lower()}||{row['vendor_name']}||{row['gl_account']}"
        if norm_key in seen_memos:
            continue
        seen_memos.add(norm_key)

        # 1. Lexical / Fuzzy similarity
        fuzzy_sim = difflib.SequenceMatcher(None, memo_clean.lower(), past_memo.lower()).ratio()

        # 2. Semantic vector similarity
        vec_sim = 0.0
        if query_emb and row["embedding"]:
            try:
                past_emb = json.loads(row["embedding"])
                vec_sim = cosine_similarity(query_emb, past_emb)
            except Exception:
                vec_sim = 0.0

        # Hybrid similarity score (giving boost to human verified precedents)
        hybrid_sim = max(vec_sim, fuzzy_sim)
        if vec_sim > 0 and fuzzy_sim > 0:
            hybrid_sim = 0.55 * vec_sim + 0.45 * fuzzy_sim

        if row["source"] == "HUMAN":
            hybrid_sim += 0.05  # slight boost for human accountant reviewed decisions

        if hybrid_sim >= min_similarity:
            scored_candidates.append({
                "txn_id": past_id,
                "memo": past_memo,
                "amount": row["amount"],
                "vendor_name": row["vendor_name"],
                "gl_account": row["gl_account"],
                "gl_description": row["gl_description"],
                "status": row["status"],
                "source": row["source"],
                "reasoning": row["reasoning"],
                "similarity": round(hybrid_sim, 3),
            })

    # Sort descending by similarity
    scored_candidates.sort(key=lambda x: x["similarity"], reverse=True)
    return scored_candidates[:top_k]


def format_rag_context_for_prompt(rag_memories: List[Dict[str, Any]]) -> str:
    """Formats retrieved RAG memories into prompt context for Amazon Nova Pro."""
    if not rag_memories:
        return ""

    lines = ["Historical Reconciliation Memory (Past Approved Precedents & Corrections):"]
    for mem in rag_memories:
        src_tag = "Human-Verified" if mem.get("source") == "HUMAN" else "Auto-Approved Precedent"
        lines.append(
            f"- Past Memo: \"{mem['memo']}\" -> Vendor: {mem['vendor_name']} "
            f"(GL {mem['gl_account']}: {mem['gl_description']}) [{src_tag}, similarity: {mem['similarity']}]"
        )
    return "\n".join(lines)


def get_audit_history_df(
    limit: int = 500,
    status_filter: Optional[str] = None,
    search_query: Optional[str] = None
) -> pd.DataFrame:
    """Retrieves audit records as a pandas DataFrame for UI inspection."""
    with get_db_connection() as conn:
        query = """
            SELECT
                id, batch_id, txn_id, txn_date, memo, amount,
                final_vendor AS vendor_name,
                final_gl_account AS gl_account,
                final_gl_description AS gl_description,
                confidence, status, source, reasoning,
                rag_sources_used, model_id, created_at, updated_at
            FROM audit_log
            WHERE 1=1
        """
        params = []
        if status_filter and status_filter != "All":
            query += " AND status = ?"
            params.append(status_filter)
        if search_query and search_query.strip():
            query += " AND (memo LIKE ? OR vendor_name LIKE ? OR gl_description LIKE ? OR txn_id LIKE ?)"
            kw = f"%{search_query.strip()}%"
            params.extend([kw, kw, kw, kw])

        query += " ORDER BY id DESC LIMIT ?"
        params.append(limit)

        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]

    if not rows:
        return pd.DataFrame(columns=[
            "id", "batch_id", "txn_id", "txn_date", "memo", "amount",
            "vendor_name", "gl_account", "gl_description",
            "confidence", "status", "source", "reasoning", "created_at"
        ])
    return pd.DataFrame(rows)


def get_audit_summary_stats() -> Dict[str, Any]:
    """Returns key metrics about audit logs and RAG memory bank."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM audit_log")
        total_records = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM audit_log WHERE status = 'AUTO_APPROVE'")
        auto_approved = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM audit_log WHERE source = 'HUMAN'")
        human_overrides = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM audit_log WHERE status IN ('REVIEW', 'MANUAL') AND source != 'HUMAN'")
        pending_review = cursor.fetchone()[0]

        # Memory precedents available for RAG
        cursor.execute("""
            SELECT COUNT(DISTINCT memo) FROM audit_log
            WHERE status IN ('AUTO_APPROVE', 'HUMAN_APPROVED', 'HUMAN_OVERRIDE')
              AND (final_vendor IS NOT NULL OR ai_vendor IS NOT NULL)
        """)
        memory_precedents = cursor.fetchone()[0]

    return {
        "total_records": total_records,
        "auto_approved": auto_approved,
        "human_overrides": human_overrides,
        "pending_review": pending_review,
        "memory_precedents": memory_precedents
    }


# Ensure DB schema is created on import
init_db()
