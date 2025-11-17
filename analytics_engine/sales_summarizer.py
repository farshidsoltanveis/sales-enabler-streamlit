# analytics_engine/sales_summarizer.py
from __future__ import annotations
import os
import time
from typing import Optional

def summarize_with_openai(
    insights: dict,
    company_name: str = "Canada Post",
    model: Optional[str] = None,
    temperature: float = 0.20,
    max_tokens: int = 350,   # hard cap
    max_sig_rows: int = 3,
    max_retries: int = 2,
    timeout_s: int = 45,
) -> str:
    """
    Produces a 5-bullet Executive Summary only, per spec.
    Requires OPENAI_API_KEY in env. Default model: gpt-4o-mini.
    """
    # <<< DELAYED IMPORT: happens only when the function is called >>>
    from analytics_to_prompt import insights_to_prompt

    from openai import OpenAI  # pip install openai>=1.0.0
    client = OpenAI(timeout=timeout_s)
    model = model or os.getenv("OPENAI_MODEL", "gpt-4o-mini")

    prompt = insights_to_prompt(
        insights,
        company_name=company_name,
        max_sig_rows=max_sig_rows,
    )

    last_err = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = client.chat.completions.create(
                model=model,
                temperature=temperature,
                max_tokens=max_tokens,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You output exactly the provided Executive Summary content—no extra sections. "
                            "Use only the numeric fields provided in the prompt. "
                            "Always state CPC vs UPS direction with mean+median gap. "
                            "When listing lanes, include lane | service | weight band + light/mid/heavy | gap%. "
                            "Never infer elasticity or revenue beyond provided numbers. Be concise and factual."
                        ),
                    },
                    {"role": "user", "content": prompt},
                ],
            )
            u = getattr(resp, "usage", None)
            if u:
                print(f"  (OpenAI usage: prompt={getattr(u,'prompt_tokens','n/a')}, "
                      f"completion={getattr(u,'completion_tokens','n/a')}, "
                      f"total={getattr(u,'total_tokens','n/a')})")
            return "### [LLM: OpenAI]\n\n" + (resp.choices[0].message.content or "").strip()
        except Exception as e:
            last_err = e
            if attempt == max_retries:
                raise
            time.sleep(1.2 * attempt)
    raise RuntimeError(f"OpenAI summarization failed after retries: {last_err}")

def summarize_offline(insights: dict, company_name: str = "Canada Post") -> str:
    """
    Offline mirror: produces the same 5-bullet Executive Summary using the prompt builder.
    """
    # <<< DELAYED IMPORT here too >>>
    from analytics_to_prompt import insights_to_prompt as _p
    text = _p(insights, company_name=company_name, max_sig_rows=3)
    return "### [LLM: Offline]\n\n" + text
