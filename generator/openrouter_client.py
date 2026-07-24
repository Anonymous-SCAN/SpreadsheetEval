"""
openrouter_client.py — thin OpenRouter chat wrapper.

Used for (1) LLM generation of business scenarios / atoms and (2) running the
frontier model as the "agent" during the difficulty-calibration eval.

Model: Kimi-k3 (moonshotai/kimi-k3) per spec, with graceful fallback to
kimi-k2 when k3 is upstream-rate-limited.
"""

from __future__ import annotations
import os
import json
import time
import threading
import urllib.request
import urllib.error

API_URL = "https://openrouter.ai/api/v1/chat/completions"

# API key is read from the environment only — never hard-code it.
#   export OPENROUTER_API_KEY=sk-or-...
API_KEY = os.environ.get("OPENROUTER_API_KEY", "")

PRIMARY_MODEL = os.environ.get("SP_MODEL", "moonshotai/kimi-k3")
FALLBACK_MODELS = ["moonshotai/kimi-k2-thinking", "moonshotai/kimi-k2"]


def chat(messages, model=None, temperature=0.7, max_tokens=8000,
         retries=5, timeout=180, no_fallback=False):
    """Call OpenRouter chat completions. Returns assistant content string.

    Tries the primary model, then falls back through FALLBACK_MODELS on
    persistent upstream rate limits / provider errors. Set no_fallback=True to
    measure a single model in isolation (used by the Pass@1 eval so a k3 timeout
    is not silently rescued by k2).
    """
    primary = model or PRIMARY_MODEL
    if no_fallback:
        models = [primary]
    else:
        models = [primary] + [m for m in FALLBACK_MODELS if m != primary]
    last_err = None
    for mdl in models:
        cur_tokens = max_tokens
        for attempt in range(retries):
            try:
                return _call_with_deadline(mdl, messages, temperature,
                                           cur_tokens, timeout)
            except _RateLimited as e:
                last_err = e
                time.sleep(min(2 ** attempt, 20))
                continue
            except _EmptyCompletion as e:
                # likely hit the token ceiling on the reasoning trace -> give it
                # more room and retry the SAME model before falling back.
                last_err = e
                cur_tokens = min(int(cur_tokens * 1.6), 60000)
                time.sleep(1.0)
                continue
            except Exception as e:  # provider/transport error -> try next model
                last_err = e
                time.sleep(1.5)
                continue
    raise RuntimeError(f"all models failed: {last_err}")


class _RateLimited(Exception):
    pass


class _EmptyCompletion(Exception):
    pass


class _Timeout(Exception):
    pass


def _call_with_deadline(model, messages, temperature, max_tokens, timeout):
    """Run _call in a worker thread with a HARD wall-clock deadline.

    urllib's socket timeout resets per read chunk, so a slow-drip streamed
    response can hang far past `timeout`. This wrapper guarantees the call
    returns (or raises _Timeout) within `timeout` seconds regardless of socket
    behaviour, so the eval never stalls on one request.
    """
    result = {}

    def worker():
        try:
            result["value"] = _call(model, messages, temperature,
                                    max_tokens, timeout)
        except BaseException as e:  # noqa: BLE001
            result["error"] = e

    th = threading.Thread(target=worker, daemon=True)
    th.start()
    th.join(timeout)
    if th.is_alive():
        raise _Timeout(f"hard deadline {timeout}s exceeded")
    if "error" in result:
        raise result["error"]
    return result["value"]


def _call(model, messages, temperature, max_tokens, timeout):
    payload = json.dumps({
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }).encode("utf-8")
    req = urllib.request.Request(
        API_URL, data=payload, method="POST",
        headers={
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "ignore")
        if e.code == 429 or "rate-limit" in body.lower():
            raise _RateLimited(body)
        raise RuntimeError(f"HTTP {e.code}: {body[:300]}")
    if "error" in data:
        msg = json.dumps(data["error"])
        if "429" in msg or "rate-limit" in msg.lower():
            raise _RateLimited(msg)
        raise RuntimeError(msg[:300])
    choice = data["choices"][0]
    content = choice["message"].get("content")
    # Reasoning models sometimes exhaust max_tokens on the reasoning trace and
    # return empty/None content with finish_reason="length". Surface that as a
    # retryable error rather than silently returning None (which downstream
    # parsing would count as a task failure — unfair to the model).
    if not content:
        fr = choice.get("finish_reason") or choice.get("native_finish_reason")
        raise _EmptyCompletion(f"empty content (finish_reason={fr})")
    return content


def extract_json(text):
    """Pull the first top-level JSON object/array out of an LLM reply."""
    import re
    # fenced block first
    m = re.search(r"```(?:json)?\s*(.*?)```", text, re.S)
    candidates = []
    if m:
        candidates.append(m.group(1))
    candidates.append(text)
    for c in candidates:
        c = c.strip()
        # find first { ... } or [ ... ] balanced
        for opener, closer in (("{", "}"), ("[", "]")):
            start = c.find(opener)
            if start == -1:
                continue
            depth = 0
            for i in range(start, len(c)):
                if c[i] == opener:
                    depth += 1
                elif c[i] == closer:
                    depth -= 1
                    if depth == 0:
                        blob = c[start:i + 1]
                        try:
                            return json.loads(blob)
                        except Exception:
                            break
    raise ValueError("no valid JSON found in LLM output")


def extract_code(text):
    """Pull the first ```python fenced code block."""
    import re
    m = re.search(r"```(?:python|py)?\s*(.*?)```", text, re.S)
    if m:
        return m.group(1).strip()
    return text.strip()


if __name__ == "__main__":
    print(chat([{"role": "user", "content": "Reply with exactly: PONG"}],
               max_tokens=10, temperature=0))
