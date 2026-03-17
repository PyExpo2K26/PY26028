"""
Ollama Client — Talks directly to local Ollama API (localhost:11434)
No internet required after initial model download.
"""
import requests
import json

OLLAMA_URL   = "http://localhost:11434/api/chat"
OLLAMA_MODEL = "qwen2.5-coder:3b"


def is_ollama_running() -> bool:
    """Check if Ollama is running locally."""
    try:
        r = requests.get("http://localhost:11434", timeout=3)
        return r.status_code == 200
    except Exception:
        return False


def is_model_available(model: str = OLLAMA_MODEL) -> bool:
    """Check if the required model is pulled."""
    try:
        r = requests.get("http://localhost:11434/api/tags", timeout=3)
        models = [m["name"] for m in r.json().get("models", [])]
        return any(model in m for m in models)
    except Exception:
        return False


def ask_llama(
    system_prompt: str,
    user_message: str,
    history: list = [],
    model: str = OLLAMA_MODEL,
    on_chunk=None
) -> str:
    """
    Send a prompt to local LLaMA via Ollama.
    Returns the full response as a string.
    """
    messages = [{"role": "system", "content": system_prompt}]
    messages += history
    messages.append({"role": "user", "content": user_message})

    payload = {
        "model": model,
        "messages": messages,
        "stream": on_chunk is not None
    }

    try:
        if on_chunk:
            with requests.post(
                OLLAMA_URL, json=payload, stream=True, timeout=120
            ) as r:
                full = ""
                for line in r.iter_lines():
                    if line:
                        chunk = json.loads(line)
                        text  = chunk.get("message", {}).get("content", "")
                        full += text
                        on_chunk(text)
                return full
        else:
            r = requests.post(
                OLLAMA_URL, json=payload,
                stream=False, timeout=120
            )
            return r.json()["message"]["content"]

    except requests.exceptions.ConnectionError:
        return "ERROR: Ollama is not running. Please start Ollama first."
    except requests.exceptions.Timeout:
        return "ERROR: Request timed out. The model may be overloaded."
    except Exception as e:
        return f"ERROR: {str(e)}"