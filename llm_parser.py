"""
llm_parser.py
Sends MD&A text to Llama 3 via Groq (free tier)
and extracts structured DCF inputs.
"""

import json
import re
import os
from groq import Groq
from dotenv import load_dotenv

load_dotenv()

client = Groq(api_key=os.getenv("GROQ_API_KEY"))

SYSTEM_PROMPT = """You are a senior investment banking analyst.
Extract forward-looking financial assumptions from MD&A text.
Return ONLY a valid JSON object. No explanation. No markdown. Just JSON."""

USER_PROMPT_TEMPLATE = """Read this MD&A excerpt and extract:
1. Revenue growth rate (%) for Year 1, Year 2, Year 3
2. EBIT margin (%)
3. Capex as % of Revenue
4. Effective tax rate (%)

Return ONLY this exact JSON format with no other text:
{{
  "revenue_growth_y1": <float>,
  "revenue_growth_y2": <float>,
  "revenue_growth_y3": <float>,
  "ebit_margin": <float>,
  "capex_pct_revenue": <float>,
  "tax_rate": <float>,
  "rationale": {{
    "revenue_growth": "<one sentence>",
    "ebit_margin": "<one sentence>",
    "capex": "<one sentence>",
    "tax_rate": "<one sentence>"
  }}
}}

MD&A Text:
{mda_text}"""

def extract_assumptions(mda_text: str) -> dict:
    """Send MD&A to Llama 3 via Groq and parse structured JSON output."""

    if not os.getenv("GROQ_API_KEY"):
        raise ValueError("GROQ_API_KEY not found. Check your .env file.")

    print("[LLM] Sending MD&A to Llama 3 via Groq...")

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": USER_PROMPT_TEMPLATE.format(
                mda_text=mda_text[:3000]
            )}
        ],
        max_tokens=512,
        temperature=0.1
    )

    raw = response.choices[0].message.content.strip()
    print(f"[LLM] Raw response received: {len(raw)} characters")

    # Extract JSON from response
    json_match = re.search(r'\{.*\}', raw, re.DOTALL)
    if not json_match:
        raise ValueError(f"LLM did not return valid JSON.\nResponse: {raw}")

    assumptions = json.loads(json_match.group())
    print("[LLM] Assumptions extracted successfully.")
    print(json.dumps(assumptions, indent=2))
    return assumptions