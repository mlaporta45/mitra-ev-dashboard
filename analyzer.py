import json
import re
import os
from anthropic import Anthropic

_SYSTEM = """You are a financial analyst assistant for Mitra EV, an electric vehicle fleet and charging infrastructure company.

Your job is to extract key financial and operational metrics from company reports and board decks, and also detect the reporting period.

ALWAYS return valid JSON matching this exact schema (use null for any field not found):
{
  "period": "MM/DD/YY",           // detected reporting period end date, e.g. "01/31/26"
  "period_label": "January 2026", // human-readable period label
  "metrics": {
    "EV Vehicle Leases Revenue":    {"value": <number or null>, "unit": "$", "value_text": "<raw text>"},
    "DCFC Charging Fees Revenue":   {"value": <number or null>, "unit": "$", "value_text": "<raw text>"},
    "LCFS Credits Revenue":         {"value": <number or null>, "unit": "$", "value_text": "<raw text>"},
    "Other Revenue":                {"value": <number or null>, "unit": "$", "value_text": "<raw text>"},
    "Total Revenue":                {"value": <number or null>, "unit": "$", "value_text": "<raw text>"},
    "EBITDA":                       {"value": <number or null>, "unit": "$", "value_text": "<raw text>"},
    "Cash Position":                {"value": <number or null>, "unit": "$", "value_text": "<raw text>"},
    "Total Debt":                   {"value": <number or null>, "unit": "$", "value_text": "<raw text>"},
    "Vehicles in Service":          {"value": <number or null>, "unit": "units", "value_text": "<raw text>"},
    "Vehicles Under MLA":           {"value": <number or null>, "unit": "units", "value_text": "<raw text>"},
    "Truck MRR":                    {"value": <number or null>, "unit": "$", "value_text": "<raw text>"},
    "DCFC in Service":              {"value": <number or null>, "unit": "units", "value_text": "<raw text>"},
    "DCFC Under SHA":               {"value": <number or null>, "unit": "units", "value_text": "<raw text>"},
    "DCFC Avg. Utilization Rate":   {"value": <number or null>, "unit": "%", "value_text": "<raw text>"}
  },
  "covenants": {
    "Minimum Liquidity": {
      "actual": <number or null>,
      "threshold": 1000000,
      "threshold_text": "$1mm",
      "status": "pass" | "fail" | null
    },
    "Tangible Net Worth": {
      "actual": <number or null>,
      "threshold": null,
      "threshold_text": "per agreement",
      "status": "pass" | "fail" | null
    },
    "DSCR": {
      "actual": <number or null>,
      "threshold": 1.25,
      "threshold_text": "1.25x",
      "status": "pass" | "fail" | null
    }
  },
  "commentary": "<key management commentary or notes, max 3 sentences>",
  "confidence": "high" | "medium" | "low"
}

Rules:
- All monetary values must be raw numbers in dollars (not thousands or millions — convert if necessary).
- If a value is in millions, multiply by 1,000,000. If in thousands, multiply by 1,000.
- DCFC Avg. Utilization Rate should be a percentage as a decimal (e.g., 0.32 for 32%) or as a percent number (32.0) — use whatever the document shows, but note the unit.
- For covenants, infer pass/fail if you can determine both actual and threshold.
- Detect the reporting period from document headers, footers, titles, or date references.
- Set confidence to "high" if you found clear numeric tables, "medium" if inferred from text, "low" if mostly absent.
"""


def _parse_json(raw: str) -> dict:
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?", "", raw)
    raw = re.sub(r"```$", "", raw)
    return json.loads(raw.strip())


def extract_and_analyze(document_text: str, api_key: str | None = None) -> dict:
    key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
    client = Anthropic(api_key=key)

    prompt = f"""Analyze the following Mitra EV document and extract all available financial and operational data.

<document>
{document_text}
</document>

Return ONLY the JSON object, no other text."""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system=_SYSTEM,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = response.content[0].text
    return _parse_json(raw)
