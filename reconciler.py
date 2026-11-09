import os
import uuid
from datetime import datetime
from typing import Optional, List, Dict, Any
import pandas as pd
from ai_engine import match_transaction
from db_audit import (
    retrieve_rag_memory,
    log_reconciliation_decision,
)
from data_loader import clean_amount_val, parse_date_sort_key




def classify_confidence(is_match: bool, confidence: float) -> str:
    if is_match and confidence >= 0.85:
        return "AUTO_APPROVE"
    if is_match and confidence >= 0.5:
        return "REVIEW"
    return "MANUAL"


def run_reconciliation(
    txns: pd.DataFrame,
    vendors: pd.DataFrame,
    batch_id: Optional[str] = None,
    use_rag: bool = True,
    progress_callback=None,
    **kwargs
) -> pd.DataFrame:
    """
    Runs AI reconciliation across all bank transactions.
    - Queries RAG memory from SQLite for similar past approved/reviewed decisions.
    - Matches using Amazon Nova Pro with reasoning and confidence scores.
    - Persists every decision into the SQLite audit log for full auditability.
    """
    if not batch_id:
        batch_id = f"batch_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"

    model_id = os.getenv("BEDROCK_MODEL_ID", "amazon.nova-pro-v1:0")
    vendor_records = vendors.to_dict("records")
    rows = []
    
    # Ensure transactions are sorted chronologically ascending by date
    txns_sorted = txns.copy()
    if "date" in txns_sorted.columns:
        txns_sorted["_sort_key"] = txns_sorted["date"].apply(parse_date_sort_key)
        txns_sorted = txns_sorted.sort_values(by=["_sort_key"]).drop(columns=["_sort_key"]).reset_index(drop=True)

    total = len(txns_sorted)

    for i, txn in enumerate(txns_sorted.to_dict("records")):
        memo = str(txn.get("memo", ""))
        amount = clean_amount_val(txn.get("amount", 0.0))
        txn_id = str(txn.get("id", f"txn_{i+1}"))
        txn_date = str(txn.get("date", datetime.utcnow().strftime("%Y-%m-%d")))

        if progress_callback:
            progress_callback(i / total, f"Analyzing ({i+1}/{total}): {memo}")

        # 1. Retrieve RAG memory from past reconciled decisions
        rag_memories: List[Dict[str, Any]] = []
        if use_rag:
            try:
                rag_memories = retrieve_rag_memory(memo, top_k=3, min_similarity=0.45)
            except Exception as e:
                print(f"[RAG Warning] Failed retrieving memory for '{memo}': {e}")
                rag_memories = []

        # 2. Invoke Bedrock Amazon Nova Pro
        try:
            ai = match_transaction(memo, amount, vendor_records, rag_memories=rag_memories)
        except Exception as e:
            ai = {
                "is_match": False,
                "confidence_score": 0.0,
                "reasoning": f"AI error: {e}",
                "vendor_name": None,
                "gl_account": None,
                "gl_description": None,
                "rag_sources_used": rag_memories,
            }

        conf_raw = ai.get("confidence_score")
        try:
            conf_val = float(conf_raw) if conf_raw is not None else 0.0
        except (ValueError, TypeError):
            conf_val = 0.0

        is_match = bool(ai.get("is_match", False))
        status = classify_confidence(is_match, conf_val)

        # 3. Log decision into SQLite audit log
        try:
            log_reconciliation_decision(
                batch_id=batch_id,
                txn_id=txn_id,
                txn_date=txn_date,
                memo=memo,
                amount=amount,
                ai_vendor=ai.get("vendor_name"),
                ai_gl_account=ai.get("gl_account"),
                ai_gl_description=ai.get("gl_description"),
                confidence=conf_val,
                status=status,
                reasoning=ai.get("reasoning", ""),
                model_id=model_id,
                rag_sources_used=rag_memories,
                source="AI",
                generate_embed=True
            )
        except Exception as e:
            print(f"[Audit Log Error] Failed logging decision for {txn_id}: {e}")

        rag_count = len(rag_memories)
        rag_label = f"🧠 {rag_count} precedent(s)" if rag_count > 0 else "None"

        rows.append({
            "id": txn_id,
            "batch_id": batch_id,
            "date": txn_date,
            "memo": memo,
            "amount": amount,
            "ai_vendor": ai.get("vendor_name"),
            "ai_gl_account": ai.get("gl_account"),
            "ai_gl_description": ai.get("gl_description"),
            "confidence": round(conf_val, 2),
            "is_match": is_match,
            "status": status,
            "rag_memory": rag_label,
            "rag_details": rag_memories,
            "reasoning": ai.get("reasoning", ""),
        })

    if progress_callback:
        progress_callback(1.0, f"Completed reconciliation for {total} transactions.")

    return pd.DataFrame(rows)