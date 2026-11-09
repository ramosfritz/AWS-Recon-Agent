import os
from typing import Optional, List, Dict, Any
import boto3
from dotenv import load_dotenv

load_dotenv()

bedrock = boto3.client(
    "bedrock-runtime",
    region_name=os.getenv("AWS_REGION", "us-east-1"),
    aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
    aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
)

MODEL_ID = os.getenv("BEDROCK_MODEL_ID", "amazon.nova-pro-v1:0")

RECON_TOOL = {
    "toolSpec": {
        "name": "submit_match",
        "description": "Submit the vendor and GL account match for a bank transaction.",
        "inputSchema": {
            "json": {
                "type": "object",
                "properties": {
                    "vendor_name": {"type": "string"},
                    "gl_account": {"type": "string"},
                    "gl_description": {"type": "string"},
                    "confidence_score": {"type": "number", "description": "0.0 to 1.0"},
                    "reasoning": {"type": "string"},
                    "is_match": {"type": "boolean", "description": "false if no good match exists"},
                },
                "required": [
                    "vendor_name", "gl_account", "confidence_score",
                    "reasoning", "is_match",
                ],
            }
        },
    }
}


def match_transaction(
    memo: str,
    amount: float,
    vendors: List[Dict[str, Any]],
    rag_memories: Optional[List[Dict[str, Any]]] = None
) -> Dict[str, Any]:
    """
    Matches bank transaction memo to vendor and GL account using Amazon Nova Pro.
    Supports RAG memory injection from past reconciled and accountant-verified transactions.
    """
    vendor_list = "\n".join(
        f"- {v.get('vendor_name', '')} (GL {v.get('gl_account', '')}: {v.get('gl_description', '')})"
        for v in vendors
    )

    rag_section = ""
    if rag_memories:
        from db_audit import format_rag_context_for_prompt
        formatted_memories = format_rag_context_for_prompt(rag_memories)
        if formatted_memories:
            rag_section = f"\n{formatted_memories}\n"

    prompt = f"""You are an expert property management accountant.

Bank transaction to reconcile:
- Memo: "{memo}"
- Amount: ${amount}

Known vendors and Chart of Accounts:
{vendor_list}
{rag_section}
Instructions:
1. Match the memo to the most likely vendor from the chart of accounts. Use fuzzy logic (e.g., "SQ *STARBUCKS" -> "Starbucks", "AMZN Mktp" -> "Amazon").
2. RAG MEMORY PRECEDENTS: If "Historical Reconciliation Memory" is provided above with prior approved decisions (especially human accountant corrections), give high weight to that precedent. If a past approved memo closely matches, adopt its vendor and GL account, and assign high confidence (>= 0.85).
3. If no reasonable match exists in known vendors or memory, set is_match to false and confidence below 0.5.
4. Always return your answer using the submit_match tool."""

    response = bedrock.converse(
        modelId=MODEL_ID,
        messages=[{"role": "user", "content": [{"text": prompt}]}],
        toolConfig={
            "tools": [RECON_TOOL],
            "toolChoice": {"tool": {"name": "submit_match"}},
        },
        inferenceConfig={"temperature": 0.0, "maxTokens": 500},
    )

    content_blocks = response.get("output", {}).get("message", {}).get("content", [])
    for block in content_blocks:
        if "toolUse" in block:
            res = block["toolUse"]["input"]
            res["rag_sources_used"] = rag_memories or []
            return res

    # Fallback if no toolUse found in response
    text_blocks = [b.get("text", "") for b in content_blocks if "text" in b]
    return {
        "is_match": False,
        "confidence_score": 0.0,
        "reasoning": f"No structured tool output returned: {' '.join(text_blocks)}",
        "vendor_name": None,
        "gl_account": None,
        "gl_description": None,
        "rag_sources_used": rag_memories or [],
    }