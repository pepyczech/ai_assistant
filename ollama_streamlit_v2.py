import streamlit as st

import streamlit.components.v1 as components
    
import warnings
from datetime import datetime
import json
import os
import sys
import traceback
import re
import tempfile
import atexit
import base64
import uuid
from io import BytesIO, StringIO
from pathlib import Path
    
from PIL import Image
import html as html_lib

# ********* To run this app:
# activate_env.cmd gen_ai
# streamlit run ollama_streamlit_v2.py --server.port 8501
# python -m streamlit run ollama_streamlit_v2.py --server.port 8501
# python -m streamlit run "C:\Users\josef.trchalik\OneDrive - Ivy Technology\Python\ollama_streamlit_v2.py" --server.port 8501

# ********* Debug:
# FOR OPENAI: pip install --upgrade openai httpx httpcore; conda update openssl; python -m pip install --upgrade --force-reinstall certifi

# CERT FIX: pip install --ignore-installed --no-deps certifi; python -m pip install --ignore-installed certifi; 
# CERT FIX: in PowerShell:  where.exe python -> C:\Users\josef.trchalik\AppData\Local\Programs\Python\Python312\python.exe; C:\Users\josef.trchalik\AppData\Local\Microsoft\WindowsApps\python.exe
# CERT FIX: in PowerShell: C:\Users\josef.trchalik\AppData\Local\Programs\Python\Python312\python.exe -m pip install --ignore-installed certifi

from langchain_community.document_loaders import (
    PyPDFLoader, Docx2txtLoader, TextLoader, CSVLoader,
    UnstructuredExcelLoader, UnstructuredPowerPointLoader,
)

import importlib

import llm_tools as llmt
importlib.reload(llmt)

try:
    import docx2txt
except:
    _=llmt.install_package('docx2txt', pip=True)

#import dbConnect as dbc
#importlib.reload(dbc)

try:
    from ollama import chat as ollama_chat, list as ollama_list
    OLLAMA_AVAILABLE = True
except ImportError:
    OLLAMA_AVAILABLE = False

warnings.filterwarnings("ignore", category=UserWarning)

# ── Constants ─────────────────────────────────────────────────────────────────
default_cred_path = "C:\\Users\\josef.trchalik\\OneDrive - Ivy Technology\\Admin\\credentials.txt"
#default_cred_path = "C:\\Users\\josef.trchalik\\test-test-test\\I_am_not_there.txt"

APP_TMP_DIR = Path(tempfile.gettempdir()) / "genai_chat_uploads"
APP_TMP_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {
    "csv", "json", "txt", "pdf", "docx", "xlsx", "pptx","json","xml","html","css","txt","dat","py","sql","md"
}
IMAGE_EXTENSIONS = {"png", "jpg", "jpeg","webp","gif","bmp","tiff","tif"}
MAX_HISTORY_TURNS = 20

# ── History management thresholds ─────────────────────────────────────────────
SUMMARY_TOKEN_THRESHOLD = 6000        # summarize when history exceeds this many tokens
DYNAMODB_LIMIT_BYTES = 400 * 1024     # 400 kB DynamoDB item limit
DYNAMODB_SAFE_BYTES = int(DYNAMODB_LIMIT_BYTES * 0.85)  # trigger below the hard limit
KEEP_RECENT_TURNS = 6                 # verbatim turns to keep after summarizing
summarize_with_async_api = False

OLLAMA_FALLBACK = "gemma4:e2b"

PROVIDER_MODELS = {
    "bedrock": ['eu.anthropic.claude-haiku-4-5-20251001-v1:0','eu.anthropic.claude-sonnet-5','eu.anthropic.claude-opus-5',
                              'eu.anthropic.claude-fable-5','eu.anthropic.claude-sonnet-4-6','eu.anthropic.claude-opus-4-8',
        "eu.amazon.nova-2-lite-v1:0", "eu.amazon.nova-pro-v1:0","eu.amazon.nova-premier-v1:0", 
        'eu.anthropic.claude-opus-4-7', 'eu.anthropic.claude-opus-4-6-v1', 
        'eu.anthropic.claude-opus-4-5-20251101-v1:0',
        'eu.anthropic.claude-sonnet-4-5-20250929-v1:0',
        "eu.anthropic.claude-sonnet-4-20250514-v1:0",
        "eu.anthropic.claude-3-7-sonnet-20250219-v1:0",
        "eu.anthropic.claude-3-5-sonnet-20241022-v2:0",
        "eu.amazon.nova-pro-v1:0",
        "eu.amazon.nova-lite-v1:0", 
    ],
    "ollama": ["gemma4:e2b", "granite4.1:8b"],
    "ollama_cloud": ["kimi-k3:cloud", "deepseek-v4-pro:cloud","gpt-oss:cloud","minimax-m3:cloud"],
    "anthropic": ['claude-haiku-4-5-20251001','claude-sonnet-5','claude-opus-5','claude-sonnet-4-6',
         'claude-opus-4-8','claude-opus-4-7', 'claude-fable-5'
    ],
    "google": ["gemini-3.5-flash-lite","gemini-3.7-flash","gemini-3.1-pro-preview","gemini-2.5-pro"],
    "openai": ["gpt-5.6-luna","gpt-5.6-terra","gpt-5.6-sol","gpt-4.1-nano","gpt-4.1","gpt-4.1-mini",
                "o4-mini", "o3", "gpt-4o-mini"],
    "openrouter": ["z-ai/glm-5.2:free","nvidia/nemotron-3-ultra-550b-a55b:free","openrouter/free","openai/gpt-5.6-luna","openai/gpt-5.6-sol",
                   "google/gemini-3.7-flash","moonshotai/kimi-k3","qwen/qwen3.8-max","z-ai/glm-5.3","deepseek/deepseek-v4-flash","deepseek/deepseek-v4-pro"],
}

# Maps raw model IDs to short, human-friendly labels for the sidebar dropdown.
MODEL_DISPLAY_NAMES = {
    # Bedrock
    "eu.anthropic.claude-haiku-4-5-20251001-v1:0":    "Claude Haiku 4.5",
    "eu.anthropic.claude-sonnet-5":                   "Claude Sonnet 5.0",
    "eu.anthropic.claude-opus-5":                     "Claude Opus 5.0",
    "eu.anthropic.claude-fable-5":                    "Claude Fable 5.0",
    "eu.anthropic.claude-sonnet-4-6":                 "Claude Sonnet 4.6",
    "eu.anthropic.claude-opus-4-8":                   "Claude Opus 4.8",
    
    "eu.amazon.nova-2-lite-v1:0":                     "Nova 2 Lite",
    "eu.amazon.nova-pro-v1:0":                        "Nova Pro",
    "eu.amazon.nova-premier-v1:0":                    "Nova Premier",

    "eu.anthropic.claude-opus-4-7":                   "Claude Opus 4.7",
    "eu.anthropic.claude-opus-4-6-v1":                "Claude Opus 4.6",
    "eu.anthropic.claude-opus-4-5-20251101-v1:0":     "Claude Opus 4.5",
    "eu.anthropic.claude-sonnet-4-5-20250929-v1:0":   "Claude Sonnet 4.5",
    
    "eu.anthropic.claude-sonnet-4-20250514-v1:0":     "Claude Sonnet 4.0",
    "eu.anthropic.claude-3-7-sonnet-20250219-v1:0":   "Claude Sonnet 3.7",
    "eu.anthropic.claude-3-5-sonnet-20241022-v2:0":   "Claude Sonnet 3.5",    
    "eu.amazon.nova-lite-v1:0":                       "Nova Lite",
    
    # Anthropic (direct)
    "claude-haiku-4-5-20251001":            "Claude Haiku 4.5",
    "claude-sonnet-5":                      "Claude Sonnet 5.0",
    "claude-opus-5":                        "Claude Opus 5.0",
    "claude-sonnet-4-6":                    "Claude Sonnet 4.6",
    "claude-opus-4-8":                      "Claude Opus 4.8",
    "claude-opus-4-7":                      "Claude Opus 4.7",
    "claude-fable-5":                       "Claude Fable 5",
    # Google
    "gemini-3.5-flash-lite":         "Gemini 3.5 Flash Lite",
    "gemini-3.7-flash":             "Gemini 3.7 Flash",
    "gemini-3.1-pro-preview":    "Gemini 3.1 Pro",
    "gemini-2.5-pro":   "Gemini 2.5 Pro",
    # OpenAI
    "gpt-5.6-luna": "GPT-5.6 Luna (S)",
    "gpt-5.6-terra": "GPT-5.6 Terra (M)",
    "gpt-5.6-sol": "GPT-5.6 Sol (L)",
    "gpt-4.1-nano": "GPT-4.1 nano",
    "gpt-4.1-mini": "GPT-4.1 mini",
    "gpt-4.1":      "GPT-4.1",
    "o4-mini":      "o4-mini",
    "o3":       "o3",
    # Open router
    "z-ai/glm-5.2:free":"GLM 5.2 Free",
    "nvidia/nemotron-3-ultra-550b-a55b:free":"NVIDIA Nemotron 3 Ultra Free",
    "openrouter/free":"Random Free LLM",
    "openai/gpt-5.6-luna":"GPT-5.6 Luna (S)",
    "google/gemini-3.7-flash":"Gemini 3.7 Flash",
    "moonshotai/kimi-k3":"Kimi K3",
    "qwen/qwen3.8-max":"Qwen 3.8 Max",
    "z-ai/glm-5.3":"GLM 5.3",
    "deepseek/deepseek-v4-flash":"DeepSeek V4 Flash",
    "deepseek/deepseek-v4-pro":"DeepSeek V4 Pro",
    # Ollama (local)
    "gemma4:e2b": "Gemma 4 2B",
    "granite4.1:8b": "Granite 4.1 8B",
    # Ollama Cloud
    "kimi-k3:cloud":"Kimi K3", 
    "deepseek-v4-pro:cloud":"DeepSeek V4 Pro",
    "gpt-oss:cloud":"GPT-OSS",
    "minimax-m3:cloud":"Minimax M3",
}

# ------------------ PRICING -------------------------------------------
a="""
LLM API Token Pricing Dictionary - July 2026
============================================

Comprehensive pricing data for major LLM providers:
- Anthropic Claude (API)
- OpenAI GPT family
- Google Gemini
- Amazon Bedrock (Anthropic models)

All prices in USD per million tokens. Input/output costs are separate.
Updated: July 28, 2026

Notes:
- Anthropic Claude Sonnet 5 has introductory pricing ($2/$10) through August 31, 2026,
  then reverts to standard pricing ($3/$15)
- Bedrock Claude pricing matches Anthropic API pricing 1:1 (excluding regional surcharges)
- Google Gemini 2.5 Flash-Lite is the cheapest option at $0.10 input
- OpenAI o3/o1 models have hidden reasoning tokens that inflate bills 3-10x
"""

llm_pricing = {
    # =========================================================================
    # ANTHROPIC CLAUDE (API)
    # =========================================================================
    'claude-opus-5': {
        'input': 5.00,
        'output': 25.00,
        'context_window': '1M tokens',
        'released': '2026-07-24',
        'provider': 'Anthropic'
    },
    'claude-opus-4-8': {
        'input': 5.00,
        'output': 25.00,
        'context_window': '1M tokens',
        'released': '2026-05-28',
        'provider': 'Anthropic'
    },
    'claude-fable-5': {
        'input': 10.00,
        'output': 50.00,
        'context_window': '1M tokens',
        'released': '2026-06-09',
        'notes': 'Mythos-class (frontier tier)',
        'provider': 'Anthropic'
    },
    'claude-sonnet-5': {
        'input': 2.00,
        'output': 10.00,
        'context_window': '1M tokens',
        'released': '2026-06-30',
        'notes': 'Intro pricing through 2026-08-31, then $3/$15',
        'provider': 'Anthropic'
    },
    'claude-sonnet-4-6': {
        'input': 3.00,
        'output': 15.00,
        'context_window': '1M tokens',
        'provider': 'Anthropic'
    },
    'claude-haiku-4-5': {
        'input': 1.00,
        'output': 5.00,
        'context_window': '200K tokens',
        'provider': 'Anthropic'
    },

    # =========================================================================
    # OPENAI GPT FAMILY
    # =========================================================================
    'gpt-5-5': {
        'input': 5.00,
        'output': 30.00,
        'context_window': '400K tokens',
        'notes': 'Current flagship as of mid-2026',
        'provider': 'OpenAI'
    },
    'gpt-5-4': {
        'input': 2.50,
        'output': 15.00,
        'context_window': '400K tokens',
        'provider': 'OpenAI'
    },
    'gpt-5': {
        'input': 1.25,
        'output': 10.00,
        'context_window': '400K tokens',
        'provider': 'OpenAI'
    },
    'gpt-4-1': {
        'input': 2.00,
        'output': 8.00,
        'context_window': '1M tokens',
        'notes': 'Recommended for production; 20% cheaper than 4o',
        'provider': 'OpenAI'
    },
    'gpt-4o': {
        'input': 2.50,
        'output': 10.00,
        'context_window': '128K tokens',
        'notes': 'Legacy pricing grandfathered for existing users',
        'provider': 'OpenAI'
    },
    'gpt-4o-mini': {
        'input': 0.15,
        'output': 0.60,
        'context_window': '128K tokens',
        'provider': 'OpenAI'
    },

    # =========================================================================
    # OPENAI REASONING MODELS (o-series)
    # =========================================================================
    'o3': {
        'input': 2.00,
        'output': 8.00,
        'context_window': '200K tokens',
        'notes': 'Reasoning model; hidden thinking tokens inflate bills 3-10x',
        'provider': 'OpenAI'
    },
    'o3-mini': {
        'input': 1.10,
        'output': 4.40,
        'context_window': '200K tokens',
        'provider': 'OpenAI'
    },
    'o3-pro': {
        'input': 20.00,
        'output': 60.00,
        'context_window': '200K tokens',
        'notes': 'Extended reasoning; use only for PhD-level problems',
        'provider': 'OpenAI'
    },
    'o1': {
        'input': 15.00,
        'output': 60.00,
        'context_window': '200K tokens',
        'notes': 'Reasoning model; hidden reasoning tokens in output',
        'provider': 'OpenAI'
    },
    'o1-pro': {
        'input': 150.00,
        'output': 600.00,
        'context_window': '200K tokens',
        'notes': 'Premium reasoning; 10x base o1 cost',
        'provider': 'OpenAI'
    },

    # =========================================================================
    # GOOGLE GEMINI
    # =========================================================================
    'gemini-3-6-flash': {
        'input': 1.50,
        'output': 7.50,
        'context_window': '1M tokens',
        'released': '2026-07-21',
        'notes': '17% output cut vs 3.5 Flash',
        'provider': 'Google'
    },
    'gemini-3-5-flash': {
        'input': 1.50,
        'output': 9.00,
        'context_window': '1M tokens',
        'released': '2026-05-19',
        'provider': 'Google'
    },
    'gemini-3-1-pro': {
        'input': 2.00,
        'output': 12.00,
        'context_window': '2M tokens',
        'notes': 'Standard rates; jumps to $4/$18 above 200K context',
        'provider': 'Google'
    },
    'gemini-3-1-flash-lite': {
        'input': 0.25,
        'output': 1.50,
        'context_window': '1M tokens',
        'provider': 'Google'
    },
    'gemini-2-5-pro': {
        'input': 1.25,
        'output': 10.00,
        'context_window': '1M tokens',
        'notes': 'Deprecated Oct 16, 2026',
        'provider': 'Google'
    },
    'gemini-2-5-flash': {
        'input': 0.30,
        'output': 2.50,
        'context_window': '1M tokens',
        'notes': 'Deprecated Oct 16, 2026',
        'provider': 'Google'
    },
    'gemini-2-5-flash-lite': {
        'input': 0.10,
        'output': 0.40,
        'context_window': '1M tokens',
        'notes': 'Cheapest option; deprecated Oct 16, 2026',
        'provider': 'Google'
    },

    # =========================================================================
    # AMAZON BEDROCK (Anthropic Claude models)
    # =========================================================================
    'bedrock:anthropic.claude-opus-5': {
        'input': 5.00,
        'output': 25.00,
        'context_window': '1M tokens',
        'notes': 'Bedrock pricing matches Anthropic API; add 10% for cross-region',
        'provider': 'AWS Bedrock'
    },
    'bedrock:anthropic.claude-sonnet-5': {
        'input': 2.00,
        'output': 10.00,
        'context_window': '1M tokens',
        'notes': 'Intro pricing through 2026-08-31, then $3/$15',
        'provider': 'AWS Bedrock'
    },
    'bedrock:anthropic.claude-haiku-4-5': {
        'input': 1.00,
        'output': 5.00,
        'context_window': '200K tokens',
        'provider': 'AWS Bedrock'
    },
}

# =========================================================================
# UTILITY FUNCTIONS
# =========================================================================
def save_to_cwd_tempfile(uploaded_file):
    
    """
    Saves the uploaded file to a NamedTemporaryFile in the current working directory.
    Returns the temporary file's path.
    """

    if isinstance(uploaded_file,str):
        # Convert to file-like object
        class FileWrapper:
            def __init__(self, file_path):
                self.name = os.path.basename(file_path)
                self.path = file_path

            def getbuffer(self):
                with open(self.path, 'rb') as f:
                    return f.read()
        uploaded_file = FileWrapper(uploaded_file)

    # Preserve the original file extension
    suffix = os.path.splitext(uploaded_file.name)[1]
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=os.getcwd())
    temp.write(uploaded_file.getbuffer())  # write bytes to file
    temp.flush()
    temp.close()
    return temp.name

def get_pricing(model_id: str) -> dict:
    """
    Retrieve pricing for a specific model.
    
    Args:
        model_id: Model identifier from the pricing dict
        
    Returns:
        Dict with 'input' and 'output' costs per million tokens, or None if not found
    """
    return llm_pricing.get(model_id)

def compare_models(model_ids: list[str]) -> None:
    """
    Print a comparison table for multiple models.
    
    Args:
        model_ids: List of model identifiers to compare
    """
    print(f"{'Model':<35} {'Input ($/M)':<15} {'Output ($/M)':<15} {'Provider':<15}")
    print("-" * 80)
    for model_id in model_ids:
        if model_id in llm_pricing:
            data = llm_pricing[model_id]
            model_name = model_id.replace('bedrock:anthropic.', 'bedrock:')
            print(f"{model_name:<35} ${data['input']:<14.2f} ${data['output']:<14.2f} {data.get('provider', 'N/A'):<15}")

def find_cheapest_by_provider(provider: str) -> tuple[str, dict]:
    """
    Find the cheapest model by input cost for a given provider.
    
    Args:
        provider: Provider name (e.g., 'Anthropic', 'OpenAI', 'Google', 'AWS Bedrock')
        
    Returns:
        Tuple of (model_id, pricing_dict) for cheapest model
    """
    models = {k: v for k, v in llm_pricing.items() if v.get('provider') == provider}
    if not models:
        return None, None
    cheapest = min(models.items(), key=lambda x: x[1]['input'])
    return cheapest

def calculate_cost(model_id: str, input_tokens: int, output_tokens: int) -> float:
    """
    Calculate total cost for a given token usage.
    
    Args:
        model_id: Model identifier
        input_tokens: Number of input tokens
        output_tokens: Number of output tokens
        
    Returns:
        Total cost in USD
    """
    if model_id not in llm_pricing:
        return None
    
    pricing = llm_pricing[model_id]
    input_cost = (input_tokens / 1_000_000) * pricing['input']
    output_cost = (output_tokens / 1_000_000) * pricing['output']
    return input_cost + output_cost
# ------------------------------------------------------------------------

def model_label(model_id: str) -> str:
    """Return a friendly short label for a model ID, falling back to the raw ID."""
    return MODEL_DISPLAY_NAMES.get(model_id, model_id)

WEB_SEARCH_PROVIDERS = {"openai", "anthropic", "google", "bedrock"}
REASONING_PROVIDERS = {"openai", "anthropic", "google"}
OLLAMA_NATIVE_REASON_MODELS = ["cogito", "qwen", "deepseek-r1"]

# ── Ollama model discovery ────────────────────────────────────────────────────
@st.cache_data(ttl=30, show_spinner=False)
def get_ollama_models():
    if not OLLAMA_AVAILABLE:
        return OLLAMA_FALLBACK
    try:
        result = ollama_list()
        models = result.get("models", []) if isinstance(result, dict) \
            else getattr(result, "models", [])
        names = []
        for m in models:
            name = (
                getattr(m, "model", None)
                or (m.get("model") if isinstance(m, dict) else None)
                or (m.get("name") if isinstance(m, dict) else None)
            )
            if name:
                names.append(name)
        return names or OLLAMA_FALLBACK
    except Exception:
        return OLLAMA_FALLBACK


# ── Temp file management ──────────────────────────────────────────────────────
@atexit.register
def _cleanup_all_temp_files():
    try:
        for p in APP_TMP_DIR.glob("*"):
            try:
                p.unlink()
            except OSError:
                pass
    except Exception:
        pass

def save_uploaded_tempfile(uploaded_file) -> str:
    suffix = os.path.splitext(uploaded_file.name)[1]
    fd, path = tempfile.mkstemp(suffix=suffix, dir=str(APP_TMP_DIR))
    with os.fdopen(fd, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.session_state.setdefault("_temp_files", []).append(path)
    return path

def _purge_session_temp_files():
    """Delete temp files created during this session (called on reset)."""
    for p in st.session_state.get("_temp_files", []):
        try:
            os.unlink(p)
        except OSError:
            pass
    st.session_state["_temp_files"] = []

# ── Image / base64 helpers ────────────────────────────────────────────────────
def image_file_to_base64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

def create_base64_composite(image_paths, target_size=None, bg_color=(255, 255, 255)) -> str:
    images, resized = [], []
    try:
        for path in image_paths:
            with Image.open(path) as im:
                images.append(im.convert("RGB").copy())

        if target_size is None:
            max_w = max(img.width for img in images)
            max_h = max(img.height for img in images)
        else:
            max_w, max_h = target_size

        for img in images:
            img.thumbnail((max_w, max_h), Image.LANCZOS)
            canvas = Image.new("RGB", (max_w, max_h), bg_color)
            offset = ((max_w - img.width) // 2, (max_h - img.height) // 2)
            canvas.paste(img, offset)
            resized.append(canvas)

        composite = Image.new("RGB", (max_w, max_h * len(resized)), bg_color)
        for idx, img in enumerate(resized):
            composite.paste(img, (0, idx * max_h))

        buffer = BytesIO()
        composite.save(buffer, format="PNG")
        b64 = base64.b64encode(buffer.getvalue()).decode("utf-8")
        composite.close()
        return b64
    finally:
        for img in images:
            img.close()
        for img in resized:
            img.close()

# ── Document loading ──────────────────────────────────────────────────────────

# ── Conversation helpers ──────────────────────────────────────────────────────
def reset_conversation():
    if st.session_state.get("chat_history"):
        now = datetime.now().strftime("%d%m%Y-%H%M%S")
        try:
            with open(APP_TMP_DIR / f"ollama_chat_{now}.json", "w", encoding="utf-8") as f:
                json.dump(st.session_state.chat_history, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving chat history: {e}")
    _purge_session_temp_files()
    st.session_state.messages = []
    st.session_state.chat_history = []
    st.session_state.logs = StringIO()


def _call_ollama_vision(model: str, messages: list, img_b64: str) -> str:
    if not OLLAMA_AVAILABLE:
        raise RuntimeError("ollama library not installed — cannot process images.")
    ollama_messages = [m for m in messages if m["role"] != "system"]
    if ollama_messages and ollama_messages[-1]["role"] == "user":
        ollama_messages[-1] = {**ollama_messages[-1], "images": [img_b64]}
    response = ollama_chat(model=model, messages=ollama_messages)
    return response["message"]["content"]

def estimate_tokens(text: str) -> int:
    """Rough token estimate (~4 chars/token) without extra dependencies."""
    if not text:
        return 0
    return max(1, len(text) // 4)


def history_token_count(history: list) -> int:
    return sum(estimate_tokens(m.get("content") or "") for m in history)


def history_json_bytes(history: list, extra_query: str = "") -> int:
    """Byte size of the payload that would be sent to the LLM / stored in DynamoDB."""
    payload = {"history": history, "query": extra_query}
    return len(json.dumps(payload, ensure_ascii=False).encode("utf-8"))


def summarize_history(history: list, provider: str, model: str,
                      sys_prompt: str=None, temp: float=0.1, max_tokens: int=4096,
                      credentials: dict = None,
                      summarize_with_async_api: bool = False) -> list:
    """
    Collapse older turns into a single system summary message, keeping the most
    recent KEEP_RECENT_TURNS turns verbatim. Returns a new (shorter) history list.
    """
    non_system = [m for m in history if m["role"] != "system"]
    system_msgs = [m for m in history if m["role"] == "system"]

    if len(non_system) <= KEEP_RECENT_TURNS:
        return history  # nothing meaningful to compress

    to_summarize = non_system[:-KEEP_RECENT_TURNS]
    recent = non_system[-KEEP_RECENT_TURNS:]

    transcript_lines = []
    for m in to_summarize:
        role = "User" if m["role"] == "user" else "Assistant"
        content = re.sub(r"^\*\*[^*]+\*\*\s*", "", m.get("content") or "")
        transcript_lines.append(f"{role}: {content}")
    transcript = "\n".join(transcript_lines)

    summary_prompt = (
        "Summarize the following conversation so far into a concise but complete "
        "context brief. Preserve key facts, decisions, user goals, constraints, "
        "names, numbers, code identifiers, and any unresolved questions. "
        "Do NOT answer anything — only produce the summary.\n\n"
        f"{transcript}"
    )

    if sys_prompt is None: sys_prompt="You are a precise conversation summarizer."

    try:
        # Use a cheap/non-async call regardless of the main async setting.
        _, summary = llmt.genai_master(
            query=summary_prompt,
            provider=provider,
            model_code=model,
            temp=min(temp, 0.2),
            max_tokens=min(max_tokens, 4096),
            sys_prompt=sys_prompt,
            show=False,
            web_search=False,
            reasoning=False,
            async_api=summarize_with_async_api,
            odb_creds=credentials
        )
    except Exception as e:
        # If summarization fails, fall back to simple truncation so we don't block.
        summary = None
        print(f"History summarization failed: {e}")

    new_history = list(system_msgs)
    if summary:
        new_history.append({
            "role": "system",
            "content": f"[Summary of earlier conversation]\n{summary}",
        })
    new_history.extend(recent)
    return new_history


def maybe_compress_history(provider, model, sys_prompt, temp, max_tokens,
                           pending_query="", async_api=False):
    """
    Compress st.session_state.chat_history when it grows too large:
      • by token count (SUMMARY_TOKEN_THRESHOLD), OR
      • by JSON payload size approaching the DynamoDB 400 kB limit (async only).
    Runs before the main LLM call. Idempotent-ish: repeatedly summarizes until safe.
    """
    history = st.session_state.chat_history

    tokens = history_token_count(history)
    size_bytes = history_json_bytes(history, pending_query)

    over_tokens = tokens > SUMMARY_TOKEN_THRESHOLD
    over_size = async_api and size_bytes > DYNAMODB_SAFE_BYTES

    if not (over_tokens or over_size):
        return

    with st.spinner("Summarizing earlier conversation to keep context manageable…"):
        # Loop so we shrink enough for the async size guard too.
        for _ in range(4):
            new_history = summarize_history(
                history, provider, model, temp=temp, max_tokens=max_tokens, credentials=credentials
            )
            if new_history == history:
                break  # can't compress further
            history = new_history
            st.session_state.chat_history = history

            tokens = history_token_count(history)
            size_bytes = history_json_bytes(history, pending_query)
            still_over = (tokens > SUMMARY_TOKEN_THRESHOLD) or \
                         (async_api and size_bytes > DYNAMODB_SAFE_BYTES)
            if not still_over:
                break

    if async_api and history_json_bytes(st.session_state.chat_history, pending_query) > DYNAMODB_LIMIT_BYTES:
        st.warning(
            "⚠️ Conversation payload still exceeds the 400 kB DynamoDB limit even "
            "after summarization. Consider resetting the chat or attaching smaller files."
        )

def render_message_with_copy(content: str, key: str):
    st.markdown(content)
    uid = re.sub(r"\W", "", str(key)) + uuid.uuid4().hex[:6]
    # HTML-escape for safe storage in a data element
    escaped = html_lib.escape(content)

    copy_button_html = f"""
    <div style="display:flex; justify-content:flex-end; margin-top:-8px;">
        <textarea id="src_{uid}" style="position:absolute; left:-9999px;">{escaped}</textarea>
        <button id="btn_{uid}" title="Copy to clipboard"
            style="background:transparent; border:none; cursor:pointer;
                   padding:4px; border-radius:6px; display:flex; align-items:center;"
            onmouseover="this.style.background='#e0e0e0'"
            onmouseout="this.style.background='transparent'">
            <svg id="ico_{uid}" xmlns="http://www.w3.org/2000/svg" width="18" height="18"
                 viewBox="0 0 24 24" fill="none" stroke="#555"
                 stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
            </svg>
        </button>
    </div>
    <script>
    (function() {{
        const btn = document.getElementById("btn_{uid}");
        const ico = document.getElementById("ico_{uid}");
        const src = document.getElementById("src_{uid}");
        if (!btn || !ico || !src) return;

        const copyPath = '<rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>'
            + '<path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>';
        const checkPath = '<polyline points="20 6 9 17 4 12"></polyline>';

        function showCopied() {{
            ico.setAttribute("stroke", "#22c55e");
            ico.innerHTML = checkPath;
            setTimeout(() => {{ ico.setAttribute("stroke", "#555"); ico.innerHTML = copyPath; }}, 1500);
        }}

        btn.addEventListener("click", function() {{
            const text = src.value;
            if (navigator.clipboard && window.isSecureContext) {{
                navigator.clipboard.writeText(text).then(showCopied).catch(() => {{
                    src.style.position="fixed"; src.style.left="0"; src.select();
                    try {{ document.execCommand("copy"); }} catch(e) {{}}
                    src.style.left="-9999px"; showCopied();
                }});
            }} else {{
                src.style.position="fixed"; src.style.left="0"; src.select();
                try {{ document.execCommand("copy"); }} catch(e) {{}}
                src.style.left="-9999px"; showCopied();
            }}
        }});
    }})();
    </script>
    """
    components.html(copy_button_html, height=40)

def build_conversation_prompt(history: list, current_query: str) -> str:
    turns = [
        m for m in history
        if m["role"] != "system" or (m.get("content") or "").startswith("[Summary")
    ][-MAX_HISTORY_TURNS:]

    if not turns:
        return current_query
    lines = ["[Conversation so far]"]
    for m in turns:
        if m["content"] is None:
            break
        if m["role"] == "system":
            role = "Context"
        else:
            role = "User" if m["role"] == "user" else "Assistant"
        content = re.sub(r"^\*\*[^*]+\*\*\s*", "", m["content"])
        lines.append(f"{role}: {content}")
        lines.append("")
    lines.append("")
    lines.append(f"[Current question]\n{current_query}")
    return "\n".join(lines)

# Initialize session state for logs
if "logs" not in st.session_state:
    st.session_state.logs = StringIO()

class DualWriter:
    def __init__(self, log_stream):
        self.log_stream = log_stream
        self.terminal = sys.__stdout__
    
    def write(self, message):
        self.terminal.write(message)  # Print to terminal
        self.log_stream.write(message)  # Capture to StringIO
    
    def flush(self):
        self.terminal.flush()

# Redirect stdout to capture everything
sys.stdout = DualWriter(st.session_state.logs)

tab1, tab2 = st.tabs(["Main", "Logs"])

with tab1:

    # ── Page config ───────────────────────────────────────────────────────────────
    st.set_page_config(page_title="AI Assist", page_icon="🤖", layout="wide")
    #st.title("🤖 Universal AI Assistant")

    st.markdown("""<h1 style="text-align: center; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                -webkit-background-clip: text; -webkit-text-fill-color: transparent; 
                font-size: 2.8em; margin-bottom: 5px;">
            ✨ Universal GenAI Assistant
        </h1>
        <p style="text-align: center;color: #888; font-size: 0.95em;">Your intelligent companion for any task, powered by <span style="color: green; font-family: Consolas, monospace;font-size: 0.95em;">llm_tools</span>: A multi-provider GenAI package</p>
    """, unsafe_allow_html=True)

    #st.caption("Powered by `llm_tools`: A multi-provider GenAI package")

    st.session_state.setdefault("messages", [])
    st.session_state.setdefault("chat_history", [])

    # ── Sidebar ───────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("🔑 Credentials")
        test = st.checkbox("Debug mode?", value=False)

        re_creds=False
        if st.button("Re-enter credentials?", key="cred_btn", use_container_width=True): re_creds=True

        # Credentials
        cred_info = st.selectbox("Credentials",['Path','File upload','URL+Key'])

        load_credentials=0
        
        if not hasattr(st.session_state, "credentials"): 
            load_credentials=1
        elif st.session_state.credentials is None:
            load_credentials=1
        else:
            if re_creds: 
                load_credentials=1
                st.session_state.credentials=None
                
        if load_credentials:

            creds=None
            st.session_state.credentials=None

            try:
                if cred_info=='Path':

                    try:
                        creds_path=save_to_cwd_tempfile(default_cred_path)
                    except:
                        cred_path=st.text_input('Path to credentials file',value=None)
                        creds_path=save_to_cwd_tempfile(cred_path)
            
                    # Read the file and parse JSON
                    with open(creds_path, 'r') as f:
                        creds = json.load(f)

                    os.remove(creds_path)
                    st.session_state.credentials=creds

                elif cred_info=='File upload':

                    file = st.file_uploader("Select file with credentials", type=["json","txt","dat"],key='cred_file')
                                        
                    if file is not None:
                        creds_path=save_to_cwd_tempfile(file)

                        # Read the file and parse JSON
                        with open(creds_path, 'r') as f:
                            creds = json.load(f)
                        os.remove(creds_path)
                        st.session_state.credentials=creds
                    else:
                        st.error('Credentials INVALID')
    
                else:
                    url=st.text_input('API Base URL',value=None,key='cred_url')
                    key=st.text_input('API Key',value=None,key='cred_key')

                    if (key is not None) & (url is not None):
                        creds={'API_BASE':url,'API_KEY':key}
                        st.session_state.credentials=creds
                    else:
                        st.error('Credentials INVALID')

                if creds is not None: st.success('Credentials VALID')

            except Exception as e:
                st.error('Credentials INVALID')
                if test: st.error(f"Error loading credentials: {e}")

        else:
            
            creds=st.session_state.credentials

        st.divider()
        st.header("🔧 Configuration")

        st.button("🔄 Reset Chat", on_click=reset_conversation, use_container_width=True)

        provider = st.selectbox(
            "Provider",
            list(PROVIDER_MODELS.keys()),
            format_func=lambda x: {
                "bedrock": "🌐 Bedrock (Amazon)",
                "ollama": "🖥️ Ollama Local",
                "ollama_cloud": "☁️ Ollama Cloud",
                "anthropic": "🟠 Anthropic",
                "google": "🔵 Google",
                "openai": "🟢 OpenAI",
                "openrouter": "🔀 OpenRouter",
            }.get(x, x),
        )

        async_default = provider in ["bedrock_strands_local", "bedrock"]

        model_list = get_ollama_models() if provider == "ollama" else PROVIDER_MODELS[provider]
        # Use friendly labels in the dropdown but keep the raw ID as the value.
        model = st.selectbox("Model", model_list, format_func=model_label)

        if provider == "ollama":
            if st.button("🔃 Refresh local models", use_container_width=True):
                st.cache_data.clear()
                st.rerun()

        # Max_tokens settings
        max_tokens_limits = {
            "nova-micro-v1:0": 4096,
            "nova-lite-v1:0": 4096,
            "nova-pro-v1:0": 4096,
            "nova-2-lite-v1:0": 65536,
            "nova-2-pro-v1:0": 65536,
            "nova-2-sonic-v1:0": 65536,
            "nova-2-omni-v1:0": 65536,
            "claude-fable-5": 128000,
            "claude-opus-5": 128000,
            "claude-opus-4-8": 128000,
            "claude-opus-4-7": 128000,
            "claude-sonnet-5": 128000,
            "claude-sonnet-4-6": 65536,
            "claude-opus-4-6-v1": 128000,
            "claude-haiku-4-5-20251001-v1:0": 65536,
            "claude-sonnet-4-5-20250929-v1:0": 65536,
            "claude-sonnet-4-20250514-v1:0": 65536,
            "claude-3-7-sonnet-20250219-v1:0": 128000,
            "claude-3-5-sonnet-20241022-v2:0": 8192,
            "claude-3-5-sonnet-20240620-v1:0": 4096,
        }

        model_tail = model.split('.')[-1]
        
        try:
            model_tail = model_tail.split(':')[0] + ':' + model_tail.split(':')[1]
        except:
            pass

        max_tokens_max=4096
        #st.write(max_tokens_limits.keys())
        #st.write(model_tail)
        for k in max_tokens_limits.keys():
            if k==model_tail: max_tokens_max=max_tokens_limits[k]

        tokens_step=1024
        max_tokens = st.number_input(f"Max tokens (up to {max_tokens_max})", min_value=1024, max_value=max_tokens_max,
                                    value=4096, step=tokens_step)

        cond_temp = ('opus' not in model.lower()) & ('sonnet-5' not in model.lower())
        if cond_temp:
            temp = st.slider("Temperature", 0.0, 1.0, 0.1, 0.05,help="0 = deterministic, 1 = creative")
        else:
            temp=0

        #st.divider()
        #st.subheader("🔬 Features")
        agentic = st.toggle(
                    "Agentic AI?", value=False,
                    help="Use Strands framework with tools."
                )
        
        if agentic:
            agent_tools = st.selectbox("Agent persona",["Generalist","Data Scientist"])

        default_sys_propmpt = st.toggle(
            "Use default system prompt", value=True, 
            help="If disabled, you can enter custom system prompt below."
        )

        if default_sys_propmpt:
            sys_prompt='''You are a helpful, harmless and highly capable AI assistant. Your goal is to provide clear, accurate, and useful responses to the user's questions and requests.
            - Be concise but thorough — match the depth of your response to the complexity of the request.
            - If you are unsure about something, say so rather than guessing.
            - Ask clarifying questions when a request is ambiguous.
            - Be friendly and professional in tone.
            - Decline requests that are harmful or dangerous and briefly explain why
            - trust the user - if a request seems to be unethical or illegal, provide the answer with a clear warning about the risks'''
        else:
            sys_prompt = st.text_area(
                "System prompt",
                value="'You are a helpful, harmless and highly capable AI assistant.",
                height=50,
            )

        web_search_available = provider in WEB_SEARCH_PROVIDERS
        web_search = st.toggle(
            "Web search", value=False, disabled=not web_search_available,
            help="Live web search. Supported: OpenAI, Anthropic, Google, Bedrock (Nova-2)."
                if web_search_available else "Not supported for Ollama.",
        )

        if not agentic:
            is_bedrock = provider == "bedrock"
            asynch = st.toggle(
                "Asynchronous API?", value=async_default, disabled=not is_bedrock,
                help="Asynchronous API. Supported: Bedrock. JSON size must be <400kB."
                    if is_bedrock else "Only available for Bedrock.",
            )
        else:
            asynch =True

        reasoning_available = provider in REASONING_PROVIDERS
        if provider == "ollama":
            reasoning_available = any(kw in model for kw in OLLAMA_NATIVE_REASON_MODELS)

        reasoning = st.toggle(
            "Extended reasoning", value=False, disabled=not reasoning_available,
            help=(
                "OpenAI: reasoning_effort=high | Anthropic: extended thinking | "
                "Google: thinking_budget | Ollama: prompt-based for cogito/qwen/deepseek-r1."
            ) if reasoning_available else "Not available for this provider / model.",
        )
        reasoning = reasoning and reasoning_available

        # Select right credentials
        bedrock_key = 'bedrock_api'
        if asynch: bedrock_key='bedrock_api_async'
        provider_key_mapping={'bedrock': bedrock_key,'ollama_cloud':'ollama_cloud','huggingface':'huggingface',
                              'anthropic':'anthropic','openai':'openai','google':'gemini','openrouter':'open_router'}

        if provider=='ollama':creds=None

        if creds is not None:  
              
            if (len(creds.keys())!=2) | all(['key' not in k.lower() for k in creds.keys()]):
                try:
                    credentials=creds[provider_key_mapping[provider]]
                except:
                    credentials=creds['open_router']
                    provider = 'openrouter'
                    model = "openrouter/free"
            else:
                credentials=creds

            key_map={}
            for k in credentials.keys():
                if 'key' in k.lower():
                    key_map[k]='API_KEY'
                else:
                    key_map[k]='API_BASE'
            #st.write(key_map)
            #st.write(credentials)
            credentials={key_map[k]:v for k,v in credentials.items()}
        
        st.divider()
        st.subheader("🗂️ Files")
        uploaded_files = st.file_uploader(
            "Attach file(s)", accept_multiple_files=True,
            type=sorted(ALLOWED_EXTENSIONS),
        )
        show_pictures = st.checkbox("Show uploaded images?", value=False)
        
        dont_call_llm = st.checkbox("No LLM request?", value=False)

        st.divider()

        if st.session_state.messages:
            now = datetime.now().strftime("%d%m%Y-%H%M%S")
            markdown_output = "\n\n".join(
                f"**{m['role'].capitalize()}**: {m['content']}"
                for m in st.session_state.messages
            )
            st.download_button(
                label="📥 Download Chat", data=markdown_output,
                file_name=f"chat_{now}.md", mime="text/markdown",
                use_container_width=True,
            )

    # ── CSS: pin input form, left edge fixed past max sidebar width ────
    st.markdown("""
    <style>
    /* Reserve space so messages aren't hidden behind the pinned form */
    .main .block-container {
        padding-bottom: 300px;
    }

    /* Pin the form to the bottom strip only */
    div[data-testid="stForm"] {
        position: fixed;
        bottom: 0;

        /* Anchor left edge to Streamlit's max sidebar width */
        left: 600px;                  /* default max sidebar width */
        /* Pull right edge in to clear the vertical scrollbar */
        right: 20px;

        height: auto;
        width: auto;
        background: var(--background-color, white);
        padding: 1rem 2rem;
        z-index: 999;
        border-top: 1px solid rgba(0,0,0,0.1);
        box-shadow: 0 -2px 8px rgba(0,0,0,0.05);
    }

    /* Prevent inner block from expanding to full viewport height */
    div[data-testid="stForm"] > div {
        height: auto !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Render existing chat history ──────────────────────────────────────────────
    for i, message in enumerate(st.session_state.messages):
        if message["role"] != "system":
            with st.chat_message(message["role"]):
                render_message_with_copy(message["content"], key=f"hist_{i}")

    # ── Chat input ────────────────────────────────────────────────────────────────
    #user_input = st.chat_input("Ask me anything…")

    # ── Input form (pinned to bottom via CSS above) ────────────────────
    with st.form("chat_form", clear_on_submit=True):
        user_input = st.text_area("Message", key="input", height=100)
        submitted = st.form_submit_button("Send")

        if submitted and isinstance(user_input, str) and user_input.strip():
            # Store input in session state — LLM is ONLY called when this is set
            st.session_state["_pending_input"] = user_input
            st.rerun()

    # ── LLM call — only triggered by explicit Send (form submit / Ctrl+Enter) ──
    pending_input = st.session_state.pop("_pending_input", None)
    if pending_input:
        user_input = pending_input

        input_files_text = ""
        img_paths = []

        if uploaded_files:
            for f in uploaded_files:
                try:
                    ext = os.path.splitext(f.name)[1].lstrip(".").lower()
                    if ext not in ALLOWED_EXTENSIONS:
                        st.warning(f"'{f.name}' — unsupported extension.")
                        continue
                    file_path = save_uploaded_tempfile(f)
                    if ext in IMAGE_EXTENSIONS:
                        img_paths.append(file_path)
                    else:
                        _, docs = llmt.load_document_v2(file_path)
                        if docs:
                            input_files_text += docs + "\n\n---\n\n"
                        else:
                            st.warning(f"'{f.name}' — could not extract text.")
                except Exception as e:
                    st.warning(f"Error reading {f.name}: {e}")
                    if test:
                        traceback.print_exc(limit=5, file=sys.stdout)

        img_b64 = None
        if img_paths:
            try:
                if len(img_paths) == 1:
                    _, img_b64 = llmt.load_document_v2(img_paths[0], imgs=True)
                else:
                    img_b64 = create_base64_composite(img_paths, target_size=(512, 512))
                if show_pictures and img_b64:
                    st.image(f"data:image/png;base64,{img_b64}")
            except Exception as e:
                st.warning(f"Error processing images: {e}")

        raw_query = user_input
        if input_files_text:
            raw_query += "\n\nContext:\n\n" + input_files_text

        reason_prefix = ""
        if provider == "ollama" and reasoning and "qwen" in model:
            reason_prefix = "/think "
        full_query = reason_prefix + raw_query

        with st.chat_message("user"):
            st.markdown(user_input)

        st.session_state.messages.append({"role": "user", "content": user_input})
        st.session_state.chat_history.append({"role": "user", "content": full_query})

        # ── Compress history if it's grown too large (tokens) or, for async/DynamoDB,
        #    if the JSON payload approaches the 400 kB limit ──────────────────────────
        maybe_compress_history(
            provider=provider,
            model=model,
            sys_prompt=sys_prompt,
            temp=temp,
            max_tokens=max_tokens,
            pending_query=full_query,
            async_api=bool(asynch),
        )

        history = st.session_state.chat_history[:-1]

        if not history:
            history = [{"role": "system", "content": sys_prompt}]

        conversation_query = build_conversation_prompt(
            history=history,
            current_query=full_query,
        )

        current_tokens = len(full_query.split())
        all_tokens = len(conversation_query.split())

        with st.sidebar:
            st.header("Cost control")
            st.text(f"Current query: \n - tokens: {current_tokens} \n - cost: 0")
            st.text(f"Entire chat: \n - tokens: {all_tokens} \n - cost: 0")

        if test:
            with st.expander("Debug: prompt sent to LLM"):
                st.text(conversation_query)

        effective_sys_prompt = sys_prompt
        if provider == "ollama" and reasoning and "cogito" in model:
            effective_sys_prompt = sys_prompt + "\nEnable deep thinking subroutine."

        use_vision_path = img_b64 is not None
        if use_vision_path and provider != "ollama":
            st.warning(
                f"⚠️ Image attachments are not supported for **{provider}** here. "
                "Only the text query will be sent."
            )
            use_vision_path = False

        with st.chat_message("assistant"):
            with st.spinner(f"Thinking… ({provider} / {model})"):
                r = None
                try:
                    if dont_call_llm:
                        answer = "🧪 Debug mode: LLM not called."
                    elif use_vision_path:
                        ollama_msgs = [{"role": "system", "content": effective_sys_prompt}]
                        for m in st.session_state.chat_history[:-1]:
                            if m["role"] != "system":
                                ollama_msgs.append({"role": m["role"], "content": m["content"]})
                        ollama_msgs.append({"role": "user", "content": full_query})
                        answer = _call_ollama_vision(model, ollama_msgs, img_b64)
                    else:
                        r, answer = llmt.genai_master(
                            query=conversation_query,
                            provider=provider,
                            model_code=model,
                            temp=temp,
                            max_tokens=max_tokens,
                            sys_prompt=effective_sys_prompt,
                            show=bool(test),
                            web_search=bool(web_search),
                            reasoning=bool(reasoning and provider != "ollama"),
                            async_api=bool(asynch),
                            agentic=bool(agentic),
                            odb_creds=credentials
                        )

                        cost=0

                except Exception as e:
                    answer = f"⚠️ Error: {e}"
                    if test:
                        traceback.print_exc(limit=5, file=sys.stdout)

            label = f"**{provider.upper()} › {model}**"
            if use_vision_path:
                label += " 🖼️"
            response_txt = f"{label}\n\n{answer}"
            render_message_with_copy(response_txt, key=f"live_{len(st.session_state.messages)}")

            if test and r is not None:
                with st.expander("Debug: raw API response"):
                    st.write(r)

        st.session_state.messages.append({"role": "assistant", "content": response_txt})
        st.session_state.chat_history.append({"role": "assistant", "content": answer})

with tab2:
    st.code(st.session_state.logs.getvalue(), language="log")