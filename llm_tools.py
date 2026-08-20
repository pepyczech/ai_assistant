from __future__ import annotations

import dataclasses
import datetime as dt
import decimal
import enum

import math

import random
from typing import Optional, List, Any, Dict, Callable

import requests

import pandas as pd
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', 200)
pd.set_option('future.no_silent_downcasting', True)

import numpy as np
import datetime
from datetime import date
today = date.today()
import time
now = datetime.datetime.now().strftime("%d%m%Y-%H%M%S")
tday = date.today().strftime("%d%m%Y")

import math
import itertools
from datetime import datetime
from difflib import SequenceMatcher
from collections import defaultdict,deque
from typing import Any, Dict, List, Tuple, Optional, Set

import subprocess, requests, sys, textwrap, json, os, sys, traceback, io, re, ast, gzip

try:
    from chromadb import Client
except:
    print('ChromaDB not installed!!!')

from bs4 import BeautifulSoup

from PIL import Image

import base64
from io import BytesIO, StringIO
# ------------------------------------------------------------------------------------ FUNCTIONS

aws=0
try:
    if "HOME" in os.environ:
        if "ec2-user" in os.environ.get("HOME"):
            aws=1
except:
    aws=0

def install_package(package_names,pip=True):

    """Installs a pip package in the current Python environment."""
    if isinstance(package_names, str):
        package_names = [package_names]
    result={}
    for package_name in package_names:
        print(f"Installing {package_name}...")
        import subprocess
        try:
            # sys.executable gets the path to the current Python interpreter
            # This runs: /path/to/python -m pip install <package_name>
            if pip:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--user', package_name])
            else:
                subprocess.check_call(['conda', 'install', '-y', '-c', 'conda-forge', package_name])
            
            print(f"Successfully installed {package_name}!")
            result[package_name]=True
        except Exception as e:
            print(f'ERROR: Failed to install package {package_name}: {e}')
            exc_type, exc_value, exc_traceback = sys.exc_info()
            traceback.print_exception(exc_type, exc_value, exc_traceback,limit=5, file=sys.stdout)
            result[package_name]=False
    
    return result

try:
    import boto3
except:
    print('WARNING: package Boto3 not installed!')

import datetime

from typing import List, Dict, Any

import tempfile

import urllib.request
import urllib.parse
from pathlib import Path

import json
import asyncio
import logging
from typing import Any, AsyncGenerator
from collections.abc import Iterable

try:
    import openai
except:
    print('WARNING: package OpenAI not installed!')

# from langchain.chat_models import ChatOpenAI # deprecated

#from langchain.document_loaders import PyPDFLoader, Docx2txtLoader, TextLoader, CSVLoader # deprecated
from langchain_community.document_loaders import PyPDFLoader, Docx2txtLoader, TextLoader, CSVLoader
#from langchain_community.document_loaders import UnstructuredExcelLoader, AzureAIDocumentIntelligenceLoader

from langchain_community.document_loaders import (
    PyPDFLoader, Docx2txtLoader, TextLoader, CSVLoader,
    UnstructuredExcelLoader, UnstructuredPowerPointLoader,
)

import charset_normalizer
from langchain_core.documents import Document

import hashlib
import ipaddress
import json
import random
import socket
import threading
import time
import unicodedata
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Any, Callable, Optional
from urllib.parse import urlparse
 
import requests

a='''
try:
    from langchain.chat_models import ChatOpenAI
except:
    from langchain_community.chat_models import ChatOpenAI

from langchain.prompts import ChatPromptTemplate

from langchain.output_parsers import ResponseSchema
from langchain.output_parsers import StructuredOutputParser

from langchain.chat_models import ChatOpenAI
from langchain.chains import ConversationChain
from langchain.memory import ConversationBufferMemory
from langchain.memory import ConversationBufferWindowMemory

from langchain.memory import ConversationTokenBufferMemory
from langchain.llms import OpenAI

from langchain.memory import ConversationSummaryBufferMemory

from langchain.chains import LLMChain

from langchain.chains import SimpleSequentialChain

from langchain.chains import SequentialChain

from langchain.chains.router import MultiPromptChain
from langchain.chains.router.llm_router import LLMRouterChain,RouterOutputParser
from langchain.prompts import PromptTemplate

from langchain.chains import RetrievalQA

from langchain_community.embeddings import HuggingFaceEmbeddings # requires latest version of sentence_transformers: pip install -U sentence-transformers

from langchain_community.llms.huggingface_pipeline import HuggingFacePipeline
from langchain.vectorstores import DocArrayInMemorySearch
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain.indexes import VectorstoreIndexCreator

'''

#from langchain.embeddings import OpenAIEmbeddings # deprecated
#from langchain_openai import OpenAIEmbeddings, OpenAI, ChatOpenAI # not on conda yet

#from sentence_transformers import SentenceTransformer, util
#from datasets import load_dataset, load_from_disk

# activate_env.cmd gen_ai
try:
    import fitz  # PyMuPDF
except:
    print('WARNING: package Fitz not installed!')

try:
    import ollama
except:
    print('WARNING: package Ollama not installed!')

# https://ai.google.dev/api/all-methods
# pip install anthropic google-generativeai

try:
    #import google-generativeai as genai
    from google import genai #pip install google-genai
except:
    print('WARNING: package Google GenAI not installed!')

try:
    import anthropic
except:
    print('WARNING: package Anthropic not installed!')

try:
    from strands import Agent, tool
    from strands.models.bedrock import BedrockModel
except Exception as e:
    print(f'Error: {e}')
    if aws:
        raise Exception('Install Strands!')
    else:
        _=install_package('strands-agents')
        from strands import Agent, tool
        from strands.models.bedrock import BedrockModel

try:
    from strands.hooks import BeforeToolCallEvent, AfterToolCallEvent
except:
    from strands.experimental.hooks import (
    BeforeToolInvocationEvent as BeforeToolCallEvent,
    AfterToolInvocationEvent as AfterToolCallEvent,
    )
    
from strands.hooks import HookProvider, HookRegistry

try:
    from strands.tools.executors import SequentialToolExecutor
except:
    print('!!! UPDATE STRANDS SDK !!!')

from strands.models.model import Model
from strands.types.content import ContentBlock, Messages, SystemContentBlock
from strands.types.streaming import StreamEvent
from strands.types.tools import ToolChoice, ToolSpec

try:
    from strands.models.ollama import OllamaModel
except Exception as e:
    print(
        "Missing Strands Ollama dependencies. Install with:\n"
        "  pip install 'strands-agents[ollama]' strands-agents-tools\n"
        "and ensure Ollama is installed and running."
    )

try:
    # Load aux DB routines
    import dbConnect as dbc
    import importlib
    importlib.reload(dbc)
except Exception as e:
    print(f'WARNING: package dbConnect not loaded: {e}')

try:
    # Load agentic tools
    import agentic_tools as agt
    import importlib
    importlib.reload(agt)
except Exception as e:
    print(f'WARNING: package agentic_tools not loaded: {e}')

urls='''
https://github.com/crewAIInc/crewAI
https://github.com/openai/openai-agents-python
https://github.com/langchain/langchain
https://github.com/Azure-Samples/python-ai-agent-frameworks-demos?utm_source=chatgpt.com
https://github.com/Significant-Gravitas/AutoGPT
https://github.com/jim-schwoebel/awesome_ai_agents?utm_source=chatgpt.com
https://github.com/NipunaRanasinghe/awesome-ai-agents?utm_source=chatgpt.com
https://github.com/slavakurilyak/awesome-ai-agents?utm_source=chatgpt.com
https://github.com/agno-agi/agno
https://github.com/microsoft/autogen
https://github.com/run-llama/smolagents
https://github.com/deepset-ai/haystack
'''

# -------------SCHEMA

def get_tools_schema():

    tools_desc = {
        "tools": [
            {
                "toolSpec": {
                    "name": "run_python",
                    "description": "Execute python source code from a string input",
                    "inputSchema": {
                        "json": {
                            "type": "object",
                            "properties": {
                                "cmd": {"type": "string"}
                            },
                            "required": ["cmd"]
                        }
                    }
                }
            },
            {
                "toolSpec": {
                    "name": "call_external_api",
                    "description": "Calls an external API based on URL using requests library.",
                    "inputSchema": {
                        "json": {
                            "type": "object",
                            "properties": {
                                "query": {"type": "string"},
                                "operator":{"type": "string"},
                                "url":{"type": "string"}
                            },
                            "required": ["query"]
                        }
                    }
                }
            },
            {
                "toolSpec": {
                    "name": "scrape_webpage",
                    "description": "Scrapes www pages using beautifulsoup library.",
                    "inputSchema": {
                        "json": {
                            "type": "object",
                            "properties": {
                                "url": {"type": "string"},
                                "css_selectors":{"type": "dict"},
                                "headers":{"type": "dict"},
                                "timeout":{"type": "float"}
                            },
                            "required": ["url","css_selectors"]
                        }
                    }
                }
            }
        ]
    }

    return tools_desc
    
# ------------ FUNCTIONS - LLM TOOLS

def run_python(cmd,show=0):
    
    fail=0
    
    try:
        out=eval(cmd)
        if isinstance(out,type(None)):
            fail=1
    except Exception as e:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        error_info=f'Error: {e}, type: {str(exc_type)}, value: {str(exc_value)}, traceback: {str(exc_traceback)}'
        if show: print(error_info)
        fail=1
        
    if fail>0:

        # Check for multi-line
        line_breaks = [';','\r\n','\n','<br>']
        line_break = None
        
        for lb in line_breaks:
            if lb in cmd:
                line_break = lb
                break

        if line_break is not None:
            lines = [c.strip() for c in cmd.split(line_break) if c.strip()]
            lines[-1] = f"print({lines[-1]})"
            cmd = "\n".join(lines)
            if show: print('Multiline command detected!')
        
        cmd=textwrap.dedent(cmd).strip()
        a=subprocess.run(
                    [sys.executable, "-c", cmd],
                    capture_output=True,
                    text=True
                )
        out=a.stdout
        if show: print(a)
        
        if (a.returncode>0) | (len(str(out))==0):
            a=subprocess.run(
                    [sys.executable, "-c", f"print('{cmd}')"],
                    capture_output=True,
                    text=True
                )
            out=a.stdout
            if show: print(a)
            
    if isinstance(out,str): out=out.replace('\n','').replace('None','')
    
    return out

def execute_oracle_sql():
    pass

def execute_vertica_sql():
    pass

def call_external_api(query,operator='search?q=',url='google.co.uk'):
    r = requests.get(f"https://{url}/{operator}{query}")
    return r.json()

def chroma_retrieve(collection=None,add_item=None,question=None):

    if isinstance(collection,type(None)):
        client = Client()
        collection = client.create_collection("agent_memory")

    if add_item:
        collection.add(documents=[add_item])

    answer=None
    if query:
        answer=collection.query(question)
    
    return(collection,answer)

def scrape_webpage(
    url: str,
    css_selectors: Dict[str, str],
    headers: Optional[Dict[str, str]] = None,
    timeout: float = 10.0
) -> Dict[str, List[str]]:
    """
    Scrape specified elements from a webpage.

    Args:
        url: The full URL of the page to scrape.
        css_selectors: A dictionary where keys are label names and values are CSS selectors
            for the data to extract. For example: {"titles": "h1.title", "links": "a.link"}.
        headers: Optional HTTP headers to include with the request (e.g., user-agent).
        timeout: How many seconds to wait for the request before failing.

    Returns:
        A dict mapping each label to a list of extracted text or attribute values.
    """
    # Default headers to mimic a browser (helps with basic bot blocks)
    headers = headers or {
        "User-Agent": "Mozilla/5.0 (compatible; AgentScraper/1.0; +https://example.com/bot)"
    }

    # Try to fetch the page
    response = requests.get(url, headers=headers, timeout=timeout)
    response.raise_for_status()  # Raises error on HTTP issues

    # Parse the HTML content
    soup = BeautifulSoup(response.content, "html.parser")
    result: Dict[str, List[str]] = {}

    # Extract using provided selectors
    for label, selector in css_selectors.items():
        items = []
        for element in soup.select(selector):
            # If scraping links, you may want attrs; otherwise use text.
            if element.name == "a" and element.has_attr("href"):
                items.append(element["href"])
            else:
                # Extract visible text, stripped of surrounding whitespace
                text = element.get_text(strip=True)
                if text:
                    items.append(text)
        result[label] = items

    return result

# ------------------------------------------------------------------------ FUNCTIONS - UTILS

def pdf_to_images(pdf_path: str, dpi: int = 300):
    """
    Convert each page of the PDF into a PIL Image (in memory)
    using PyMuPDF, no Poppler.
    """
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        # Render page to pixmap
        pix = page.get_pixmap(dpi=dpi)
        # Get PNG bytes
        img_bytes = pix.tobytes("png")
        image = Image.open(io.BytesIO(img_bytes))
        images.append(image)
    return images

def image_to_tempfile(image: Image.Image, dirname: str = "temp_images"):
    """
    Save a PIL Image to a temporary file, return the path.
    Ollama Python client can accept image paths as input images.
    """
    os.makedirs(dirname, exist_ok=True)
    temp_path = os.path.join(dirname, f"page_{os.urandom(8).hex()}.png")
    image.save(temp_path, format="PNG")
    return temp_path

def pdf_scan_to_text(pdf_path: str,
                     provider: str = "aws",
                     model_code: str = 'eu.anthropic.claude-sonnet-4-6',
                     prompt_template: str = "Extract all readable text from the data, preserving the layout. Return only the extracted text with no commentary.",
                     dpi: int = 150,
                     temp: float = 0.0,
                     max_tokens: int = 80000,
                     odb_creds=None,
                     async_api: int = 1,
                     img_format: str = "jpeg",
                     show: int = 0):
    
    """
    Convert PDF pages to images and send them to a vision-capable LLM via
    genai_master() to perform OCR. Returns concatenated results from all pages.

    Parameters
    ----------
    pdf_path : str
        Path to the source PDF.
    provider : str
        genai_master provider. Must be a VISION-capable model/provider.
        - 'aws'/'bedrock'   -> uses Bedrock Converse image content blocks
                               (recommended: a Claude 3.x / Sonnet vision model)
        - 'anthropic'/'claude', 'openai'/'gpt', 'google'/'gemini' -> handled
                               via provider-specific multimodal payloads below.
    model_code : str
        Vision-capable model id. Examples:
        - Bedrock:   'eu.anthropic.claude-3-5-sonnet-20241022-v2:0'
        - Anthropic: 'claude-sonnet-4-6'
        - OpenAI:    'gpt-4o'
        - Gemini:    'gemini-2.5-flash'
    prompt_template : str
        Instruction sent alongside each page image.
    dpi : int
        Render resolution for pdf_to_images() - for async API, there's 400kB size limit for all DynamoDB items: 
        At 150 DPI JPEG a typical page is well under 400 KB. This helps regardless of which API path you use.
    temp : float
        Sampling temperature (0.0 recommended for OCR fidelity).
    max_tokens : int
        Max output tokens per page.
    odb_creds : dict | str | None
        Credentials passed straight through to genai_master().
    async_api : int
        Passed through to genai_master() (Bedrock async proxy).
    img_format : str
        'png' or 'jpeg'. Controls how the rendered page is encoded.
    show : int
        Verbosity flag forwarded to genai_master().

    Returns
    -------
    (pages, full_text)
        pages     : list of {'page_content': <str>} dicts (langchain-like)
        full_text : concatenated OCR text across all pages
    """
    img_format = img_format.lower()
    if img_format == "jpg":
        img_format = "jpeg"
    mime = f"image/{img_format}"

    OCR_SYS_PROMPT = ("You are an OCR engine. Transcribe the text from the provided "
                      "image exactly, preserving layout where sensible. Do not add commentary.")

    images = pdf_to_images(pdf_path, dpi=dpi)
    if show:
        print(f"[pdf_scan_to_text] {len(images)} page(s) rendered from {pdf_path}")

    def _encode_image(pil_img):
        """PIL image -> (raw_bytes, base64_str)."""
        save_fmt = "JPEG" if img_format == "jpeg" else "PNG"
        # JPEG has no alpha channel
        if save_fmt == "JPEG" and pil_img.mode in ("RGBA", "P", "LA"):
            pil_img = pil_img.convert("RGB")
        buf = io.BytesIO()
        pil_img.save(buf, format=save_fmt)
        raw = buf.getvalue()
        return raw, base64.b64encode(raw).decode("utf-8")

    def _build_query(pil_img):
        raw_bytes, b64 = _encode_image(pil_img)
        p = provider.lower()

        # Bedrock Converse expects raw bytes, not base64
        if any(t in p for t in ("bedrock", "aws", "amazon", "nova")):
            # On real AWS -> boto3 converse needs raw bytes.
            # Off AWS -> goes over JSON HTTP API, needs base64 string.
            on_aws = "HOME" in os.environ and "ec2-user" in os.environ.get("HOME", "")
            image_source = {"bytes": raw_bytes} if on_aws else {"bytes": b64}
            return [{
                "role": "user",
                "content": [
                    {"text": prompt_template},
                    {"image": {"format": img_format, "source": image_source}},
                ],
            }]

        if any(t in p for t in ("anthropic", "claude", "sonnet", "opus", "haiku")):
            return [
                {"type": "text", "text": prompt_template},
                {"type": "image",
                 "source": {"type": "base64", "media_type": mime, "data": b64}},
            ]

        if any(t in p for t in ("openai", "gpt", "chatgpt")):
            return [
                {"type": "text", "text": prompt_template},
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
            ]

        if any(t in p for t in ("google", "gemini", "vertex")):
            return [
                {"text": prompt_template},
                {"inline_data": {"mime_type": mime, "data": b64}},
            ]

        raise ValueError(
            f"Provider '{provider}' is not configured for vision OCR. Use a "
            "Bedrock/Anthropic/OpenAI/Gemini vision-capable provider & model."
        )

    page_texts, pages = [], []
    total = len(images)

    for idx, img in enumerate(images):

        query = _build_query(img)

        if show:
            print(f"[pdf_scan_to_text] OCR page {idx + 1}/{total} via {provider} / {model_code}")
            print(f"********** Query page {idx}: type={type(query)}, blocks={len(query)}")
            with open(f"query_page_{idx}.txt", "w", encoding="utf-8") as f:
                f.write(str(query))

        #print(pepa)

        try:
            _, answer = genai_master(
                query,
                provider=provider,
                model_code=model_code,
                temp=temp,
                max_tokens=max_tokens,
                odb_creds=odb_creds,
                show=show,
                web_search=0,
                sys_prompt=OCR_SYS_PROMPT,
                async_api=async_api,
            )
            text = answer if isinstance(answer, str) else str(answer)
        except Exception as e:
            text = f"[OCR ERROR on page {idx + 1}: {e}]"
            if show:
                print(text)

        page_texts.append(text)
        pages.append({"page_content": text})

    return pages, "\n\n".join(page_texts)

def pdf_scan_to_text_ollama(pdf_path: str,
                            model_name: str = "gemma4:e4b",
                            prompt_template: str = "Extract all readable text from this image, preserving the layout.",
                            dpi: int = 300):
    """OCR a PDF via a local Ollama vision model. Returns (pages, full_text)."""
    images = pdf_to_images(pdf_path, dpi=dpi)
    image_paths = [image_to_tempfile(img) for img in images]

    page_texts, pages = [], []
    try:
        for idx, img_path in enumerate(image_paths):
            try:
                resp = ollama.chat(
                    model=model_name,
                    messages=[{"role": "user",
                               "content": prompt_template,
                               "images": [img_path]}],
                    stream=False,
                )
                text = resp.get("message", {}).get("content", "")
            except Exception as e:
                text = f"[OCR ERROR on page {idx + 1}: {e}]"
            page_texts.append(text)
            pages.append({"page_content": text})
    finally:
        # image_to_tempfile creates files that should be cleaned up
        for path in image_paths:
            try:
                os.remove(path)
            except OSError:
                pass

    return pages, "\n\n".join(page_texts)

# Loaders keyed by extension keeps the dispatch readable and extensible
def load_document(file, imgs=False, test=0):
    """
    Load a document into (data, string_data).

    Returns:
        data        : list[Document] for text docs, or OCR result, or '' on failure
        string_data : concatenated text content (or base64 string for images)

    NOTE: `data` type is not uniform across branches (historical behaviour
    preserved). Callers should generally rely on `string_data`.
    """
    _, extension = os.path.splitext(file)
    extension = extension.lower()
    print(f"Loading {file}")

    # --- Image files: return base64 directly, no loader ---
    if extension in (".png", ".jpg", ".jpeg", ".jpe", ".jp2", ".jpx"):
        print("This is an image")
        with open(file, "rb") as f:
            string_data = base64.b64encode(f.read()).decode("utf-8")
        # base64 output is already correctly padded — no manual padding needed
        return None, string_data

    # --- Choose loader by extension ---
    if extension == ".pdf":
        loader = PyPDFLoader(file_path=file, extract_images=imgs)
    elif extension == ".docx":
        loader = Docx2txtLoader(file)
    elif extension in (".ppt", ".pptx"):
        loader = UnstructuredPowerPointLoader(file, mode="elements")
    elif extension == ".txt":
        loader = TextLoader(file)
    elif extension == ".csv":
        loader = CSVLoader(file_path=file)
    elif extension in (".xls", ".xlsx", ".xlsm"):  # requires `unstructured`
        loader = UnstructuredExcelLoader(file, mode="elements")
    else:
        print("Document format is not supported!")
        return None, None

    data = loader.load()
    docs_text = [doc.page_content for doc in data]
    total_len = sum(len(t) for t in docs_text)

    if total_len > 0:
        return data, "\n\n---\n\n".join(docs_text)

    # Empty extraction from a PDF => likely a scanned/image-based PDF
    if extension == ".pdf":
        print(f"WARNING: file {file} is image-based, running OCR...")
        return pdf_scan_to_text(file, show=test,dpi=150,img_format='jpeg')

    return "", ""

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

def load_text_file_safely(file_path: str) -> list[Document]:
    with open(file_path, "rb") as f:
        raw = f.read()
    result = charset_normalizer.from_bytes(raw).best()
    if result is None:
        with open(file_path, "r", encoding="latin-1") as f:
            text = f.read()
    text = str(result)
    return [Document(page_content=text, metadata={"source": file_path})]

def load_document_v2(file_path: str, imgs: bool = False):
    _, extension = os.path.splitext(file_path)
    extension = extension.lower()
    success=1
    if extension == ".pdf":
        loader = PyPDFLoader(file_path=file_path, extract_images=imgs)
    elif extension == ".docx":
        loader = Docx2txtLoader(file_path)
    elif extension in (".ppt", ".pptx"):
        loader = UnstructuredPowerPointLoader(file_path, mode="elements")
    elif extension in (".txt",".dat",".md"):
        loader = TextLoader(file_path, autodetect_encoding=True)
    elif extension == ".csv":
        loader = CSVLoader(file_path=file_path)
    elif ".xls" in extension:
        loader = UnstructuredExcelLoader(file_path, mode="elements")
    elif extension in (".png", ".jpg", ".jpeg", ".jpe", ".jp2", ".jpx"):
        text = image_file_to_base64(file_path)
        docs=[text]
    else:
        success=0
        docs=None
        text=None

    if success:
        try:
            docs = loader.load()
        except:
            docs=load_text_file_safely(file_path)
        text = "\n\n---\n\n".join(doc.page_content for doc in docs)
    elif extension==".json":
        with open(file_path, "r", encoding="utf-8") as f:
            text = json.dumps(json.load(f))
        docs = [text]
    elif extension==".xml":
        try:
            import xmltodict
        except:
            _=install_package('xmltodict')
            import xmltodict
        with open(file_path, 'r', encoding='latin-1') as f:
            text = json.dumps(xmltodict.parse(f.read()))
        docs = [text]
    elif any([extension==ext for ext in ['.html','.css','.txt','.dat','.py','.sql']]):
        with open(file_path, "r", encoding="latin-1") as f:
            text = f.read()
        docs = [text]
        
    return docs, text

# ----------------------------------------------------------------------- FUNCTIONS - CALLING LLMS

# NEW FUNCTION

def call_bedrock_aws(prompt=None,model_id='us.amazon.nova-lite-v1:0', system_prompt='You are a helpful assistant',
                     max_tokens=1024, temp=0.1, tools_desc=None, show=0):
    '''
   MODELS
    'us.meta.llama3-2-11b-instruct-v1:0'
    
    'eu.amazon.nova-lite-v1:0' - WORKS US & EU
    'eu.amazon.nova-pro-v1:0' - WORKS US & EU

    'us.amazon.nova-2-lite-v1:0' - WORKS, US ONLY
    
    'anthropic.claude-fable-5'
    'anthropic.claude-opus-4-8'
    'us.anthropic.claude-opus-4-7'
    'us.anthropic.claude-sonnet-4-6'
    'us.anthropic.claude-opus-4-6-v1'
    'us.anthropic.claude-haiku-4-5-v1:0'
    'us.anthropic.claude-sonnet-4-5-20250929-v1:0'
    'eu.anthropic.claude-sonnet-4-20250514-v1:0' - WORKS US & EU
    'eu.anthropic.claude-3-7-sonnet-20250219-v1:0' - WORKS  US & EU
    'eu.anthropic.claude-3-5-sonnet-20241022-v2:0' - WORKS US & EU
    
    'eu.anthropic.claude-3-5-sonnet-20240620-v1:0' - WORKS US & EU

    # ********** PARAMETERS
    temp - low -> deterministic, high -> creative
    max_tokens - <=0 triggers max output token mapping routine

    TOOLS_DESC - No tools: None, otherwise dictionary in standard Bedrock format, e.g.:
    tools_desc= {
            "tools": [
                {
                    "toolSpec": {
                        "name": "tavily_search_tool",
                        "description": "Searches Tavily for research topics.",
                        "inputSchema": {
                            "json": {
                                "type": "object",
                                "properties": {
                                    "query": {"type": "string"}
                                },
                                "required": ["query"]
                            }
                        }
                    }
                },
                {
                    "toolSpec": {
                        "name": "arxiv_search_tool",
                        "description": "Searches arXiv for academic papers.",
                        "inputSchema": {
                            "json": {
                                "type": "object",
                                "properties": {
                                    "query": {"type": "string"}
                                },
                                "required": ["query"]
                            }
                        }
                    }
                }
            ]
        }
    
    # Test
    a=call_bedrock_aws(prompt=None,model='eu.amazon.nova-lite-v1:0', show=1)
    
    '''
    
    try:
        import json, sys
        import boto3
        from botocore.config import Config
        
        # Longer timeouts for long generations
        config = Config(
            connect_timeout=120,   # seconds to wait for connection
            read_timeout=3600      # seconds to wait for full response
        )

        # Max_tokens settings
        max_tokens_limits = {
            "nova-micro-v1:0": 4096,
            "nova-lite-v1:0": 4096,
            "nova-pro-v1:0": 4096,
            "nova-2-lite-v1:0": 65536,
            "nova-2-pro-v1:0": 65536,
            "nova-2-sonic-v1:0": 65536,
            "nova-2-omni-v1:0": 65536,
            "claude-opus-5": 128000,
            "claude-opus-4-8": 128000,
            "claude-opus-4-7": 128000,
            "claude-sonnet-5": 128000,
            "claude-sonnet-4-6": 65536,
            "claude-opus-4-6-v1": 128000,
            "claude-haiku-4-5-v1:0": 65536,
            "claude-sonnet-4-5-20250929-v1:0": 65536,
            "claude-sonnet-4-20250514-v1:0": 65536,
            "claude-3-7-sonnet-20250219-v1:0": 128000,
            "claude-3-5-sonnet-20241022-v2:0": 8192,
            "claude-3-5-sonnet-20240620-v1:0": 4096,
        }
        
        if prompt is None: 
            prompt= [
                        {
                            "role": "user",
                            "content": [
                                {"text": 'Explain the differences and similarities between reinforcement learning and transfer learning.'}
                            ]
                        }
                    ]
            
        elif isinstance(prompt,str):
            try:
                prompt=json.loads(prompt)
                if isinstance(prompt,dict): prompt=[prompt]
            except:
                prompt= [
                            {
                                "role": "user",
                                "content": [
                                    {"text": prompt}
                                ]
                            }
                        ]
    
        if model_id[0:2]!='us': 
            region="eu-central-1"
        else:
            region="us-east-1"
        
        if str(tools_desc).lower() in ['none','','{}']: tools_desc=None

        if isinstance(tools_desc,str):
            try:
                tools_desc=json.loads(tools_desc)
            except:
                tools_desc=None
        
        # If the caller sent max_tokens <= 0, pick a model-specific practical cap.
        if max_tokens <= 0:
            # Find the matching family tail (e.g., "...sonnet-4-20250514-v1:0")
            model_tail = model_id.split('.', 1)[-1]
            max_tokens = next(
                (v for k, v in max_tokens_limits.items() if k == model_tail.split(':')[0] + ':' + model_tail.split(':')[1]),
                1024
            )

        # Normalize modelId prefix to region family (`us.` or `eu.`), then force to desired region
        if model_id[:3] not in ['us.', 'eu.']:
            model_id = 'eu.' + model_id
        if region[:2] != model_id[:2]:
            model_id = region[:2] + model_id[2:]

        if show: print(f'************** REGION: {region}')

        bedrock = boto3.client(
            'bedrock-runtime',
            config=config,
            region_name=region
        )

        # Build a unified Converse payload (works across providers)
        conversation = prompt

        # Common inference parameters supported by Converse
        # maxTokens, temperature, topP, stopSequences are portable across models.
        cond_temp = ('opus' not in model_id.lower()) & ('sonnet-5' not in model_id.lower())
        if cond_temp:
            inference_cfg = {
                "temperature": float(temp),
                "stopSequences": [],
                "maxTokens": int(max_tokens)
            }
        else:
            # For models where temperature is depreciated
            inference_cfg = {
                "stopSequences": [],
                "maxTokens": int(max_tokens)
            }

        # Optional: per‑model unique fields can go in additionalModelRequestFields
        # (left empty here for simplicity)

        # Converse API request & response shape. [1](https://boto3.amazonaws.com/v1/documentation/api/latest/reference/services/bedrock-runtime/client/converse.html)
        #                                        [2](https://docs.aws.amazon.com/bedrock/latest/APIReference/API_runtime_Converse.html)
        
        if tools_desc is not None:
            if show: print(f'>>>> using tools: {tools_desc}')
            response = bedrock.converse(
                modelId=model_id,
                system=[{"text": system_prompt}],
                messages=conversation,
                toolConfig=tools_desc,
                inferenceConfig=inference_cfg
                # additionalModelRequestFields={},       # if a model requires extra vendor-specific params
                # guardrailConfig=...,   # add as needed
            )

        else:
            if show: print('>>>> not using tools')
            response = bedrock.converse(
                modelId=model_id,
                system=[{"text": system_prompt}],
                messages=conversation,
                inferenceConfig=inference_cfg
            )

        if show: print(f'ENTIRE RESPONSE: {response}')
            
        # Extract first text block from the assistant message
        output_message = response.get("output", {}).get("message", {})

        '''
        content_blocks = output_message.get("content", []) or []
        result_text = ""
        for block in content_blocks:
            if "text" in block and isinstance(block["text"], str):
                result_text = block["text"]
                break
        '''
        if show:
            print(f'OUTPUT \n {output_message}')
            print(f'PROMPT \n {prompt}')
            
        return(output_message,prompt)

    except Exception as e:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        error_info=f'Error: {e}, type: {str(exc_type)}, value: {str(exc_value)}, traceback: {str(exc_traceback)}'
        if show:
            print(f'ERROR: {error_info}')
            print(f'PROMPT \n {prompt}')
        return(error_info, prompt)

# NEW FUNCTION
# AWS LAMBDA: arn:aws:lambda:eu-central-1:730335465043:function:aws-bedrock-api-v0-BedrockLambdaFunction-RjV60Y6J1th3
# APP: aws-bedrock-api-v0
# URL: 'https://gzukpgh4eg.execute-api.eu-central-1.amazonaws.com/prod/invoke'

def call_bedrock_api(prompt=None,model='us.amazon.nova-lite-v1:0', api_url=None, api_key=None, system_prompt='You are a helpful assistant',
                     max_tokens=1024, temp=0.1, tools_desc=None, show=0):

    '''
    import requests

    AWS LAMBDA: arn:aws:lambda:eu-central-1:730335465043:function:aws-bedrock-api-v0-BedrockLambdaFunction-RjV60Y6J1th3
    APP: aws-bedrock-api-v0
    
    MODELS
    'us.meta.llama3-2-11b-instruct-v1:0'
    
    'eu.amazon.nova-lite-v1:0' - WORKS US & EU
    'eu.amazon.nova-pro-v1:0' - WORKS US & EU

    'us.amazon.nova-2-lite-v1:0' - WORKS, US ONLY
    
    'anthropic.claude-fable-5'
    'anthropic.claude-opus-4-8'
    'us.anthropic.claude-opus-4-7'
    'us.anthropic.claude-sonnet-4-6'
    'us.anthropic.claude-opus-4-6-v1'
    'us.anthropic.claude-haiku-4-5-v1:0'
    'us.anthropic.claude-sonnet-4-5-20250929-v1:0'
    'eu.anthropic.claude-sonnet-4-20250514-v1:0' - WORKS US & EU
    'eu.anthropic.claude-3-7-sonnet-20250219-v1:0' - WORKS  US & EU
    'eu.anthropic.claude-3-5-sonnet-20241022-v2:0' - WORKS US & EU
    
    'eu.anthropic.claude-3-5-sonnet-20240620-v1:0' - WORKS US & EU

    SYNC: 'https://gzukpgh4eg.execute-api.eu-central-1.amazonaws.com/prod/invoke'
    ASYNC: 'https://qt2gq77oik.execute-api.eu-central-1.amazonaws.com/prod/invoke'

    # ********** PARAMETERS
    temp - low -> deterministic, high -> creative
    max_tokens - <=0 triggers max output token mapping routine

    TOOLS_DESC - No tools: None, otherwise dictionary in standard Bedrock format, e.g.:
    tools_desc= {
            "tools": [
                {
                    "toolSpec": {
                        "name": "tavily_search_tool",
                        "description": "Searches Tavily for research topics.",
                        "inputSchema": {
                            "json": {
                                "type": "object",
                                "properties": {
                                    "query": {"type": "string"}
                                },
                                "required": ["query"]
                            }
                        }
                    }
                },
                {
                    "toolSpec": {
                        "name": "arxiv_search_tool",
                        "description": "Searches arXiv for academic papers.",
                        "inputSchema": {
                            "json": {
                                "type": "object",
                                "properties": {
                                    "query": {"type": "string"}
                                },
                                "required": ["query"]
                            }
                        }
                    }
                }
            ]
        }
    
    # Test
    o,p=call_bedrock_api(prompt=None,model='eu.amazon.nova-lite-v1:0', api_url=api_url, api_key=api_key, show=1)
    
    '''
    
    if prompt is None: 
        prompt= [
                    {
                        "role": "user",
                        "content": [
                            {"text": 'Explain the differences and similarities between reinforcement learning and transfer learning.'}
                        ]
                    }
                ]
        
    elif isinstance(prompt,str):
        try:
            prompt=json.loads(prompt)
            if isinstance(prompt,dict): prompt=[prompt]
        except:
            prompt= [
                        {
                            "role": "user",
                            "content": [
                                {"text": prompt}
                            ]
                        }
                    ]
            
    if str(tools_desc).lower() in ['none','','{}']: tools_desc=None
    if isinstance(tools_desc,str):
        try:
            tools_desc=json.loads(tools_desc)
        except:
            tools_desc=None
            
    if (api_url is None) | (api_key is None):
        try:
            odb_creds=dbc.jsonpass(pattern=None,fn='..//Admin//credentials.txt')
            api_url =odb_creds['bedrock_api']['API_BASE']
            api_key =odb_creds['bedrock_api']['API_KEY']
        except:
            raise Exception('Cannot find suitable credentials...')

    if model[0:2]!='us': 
        region="eu-central-1"
    else:
        region="us-east-1"
    
    if show: print(f'************** REGION: {region}')

    # Your prompt or input to the GenAI model
    payload = {
        "prompt": prompt,
        "region": region,
        "modelId": model,
        "system_prompt":system_prompt,
        "max_tokens": max_tokens,
        "temp": temp,
        "tools_desc":tools_desc
    }
    
    # Headers including the API key
    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key
    }
    
    # Make the POST request
    response = requests.post(api_url, json=payload, headers=headers)
    
    # Print the response
    if response.status_code == 200:
        r=response.json()
        
        if show: print("Response:", response.json())
        out=response.json()
    else:
        if show: print("Error:", response.status_code, response.text)
        out=f'ERROR - Status code: {response.status_code}; {response.text}'

    return(out,prompt)

def call_bedrock_api_async(prompt=None,model='us.amazon.nova-lite-v1:0', api_url=None, api_key=None, system_prompt='You are a helpful assistant',
                     max_tokens=1024, tools_desc = None, temp=0.1, show=0):

    '''
    import requests

    prompt = payload.get("prompt", dummy_prompt)
    region = payload.get("region", DEFAULT_REGION)
    model_id = payload.get("modelId", DEFAULT_MODEL_ID)
    system = payload.get("system_prompt", DEFAULT_SYSTEM)
    temp = float(payload.get("temp", DEFAULT_TEMP))
    max_tokens = int(payload.get("max_tokens", DEFAULT_MAX_TOKENS_SAFE))
    tools_desc = payload.get("tools_desc", None)

    
    AWS LAMBDA: arn:aws:lambda:eu-central-1:730335465043:function:aws-bedrock-api-v0-BedrockLambdaFunction-RjV60Y6J1th3
    APP: aws-bedrock-api-v0
    
    MODELS
    'us.meta.llama3-2-11b-instruct-v1:0'
    
    'us.amazon.nova-lite-v1:0' - WORKS
    'us.amazon.nova-pro-v1:0' - WORKS

    'us.amazon.nova-2-lite-v1:0' - WORKS, US ONLY
    
    'us.anthropic.claude-sonnet-4-5-20250929-v1:0'
    'us.anthropic.claude-sonnet-4-20250514-v1:0' - WORKS
    'us.anthropic.claude-3-7-sonnet-20250219-v1:0' - WORKS
    'us.anthropic.claude-3-5-sonnet-20241022-v2:0' - WORKS
    
    'us.anthropic.claude-3-5-sonnet-20240620-v1:0' - WORKS

    # ********** PARAMETERS
    temp - low -> deterministic, high -> creative
    max_tokens - <=0 triggers max output token mapping routine
    
    # Test
    a=call_bedrock_api(prompt=None,model='eu.amazon.nova-lite-v1:0', api_url=api_url, api_key=api_key, show=1)

    SYNC: 'https://gzukpgh4eg.execute-api.eu-central-1.amazonaws.com/prod/invoke'
    ASYNC: 'https://qt2gq77oik.execute-api.eu-central-1.amazonaws.com/prod/invoke'
    
    '''
    
    if prompt is None: 
        prompt= [
                    {
                        "role": "user",
                        "content": [
                            {"text": 'Explain the differences and similarities between reinforcement learning and transfer learning.'}
                        ]
                    }
                ]
            
    if str(tools_desc).lower() in ['none','','{}']: tools_desc=None
    if isinstance(tools_desc,str):
        try:
            tools_desc=json.loads(tools_desc)
        except:
            tools_desc=None
    
    if (api_url is None) | (api_key is None):
        try:
            #import dbConnect as dbc
            odb_creds=dbc.jsonpass(pattern=None,fn='..//Admin//credentials.txt')
            api_url =odb_creds['bedrock_api_async']['API_BASE']
            api_key =odb_creds['bedrock_api_async']['API_KEY']
        except:
            raise Exception('Cannot find suitable credentials...')

    if model[0:2]!='us': 
        region="eu-central-1"
    else:
        region="us-east-1"

    if isinstance(prompt,str):
        prompt=[{"role": "user", "content": [{"text": prompt}]}]
    elif isinstance(prompt,list) & ('role' in prompt[0].keys()) & ('content' in prompt[0].keys()):
        pass
    else:
        raise Exception(f'!!! Invalid prompt data type or format. \n \n Type: {type(prompt)} \n \n Format: {prompt}')
        
    #api_url = "https://qt2gq77oik.execute-api.eu-central-1.amazonaws.com/prod/invoke"
    
    # Your prompt or input to the GenAI model
    
    payload = {
        "prompt": prompt,
        "region": region,
        "modelId": model,
        "system_prompt":system_prompt,
        "max_tokens": max_tokens,
        "temp": temp,
        "tools_desc":tools_desc
    }
    
    # Headers including the API key
    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key
    }

    if show:
        print(f'PAYLOAD: {payload}')
        print(f'URL: {api_url}')
    
    # ************ Initial (POST) request
    
    try:
        response = requests.post(
            api_url,
            headers=headers,
            data=json.dumps(payload),
            timeout=10
        )
        
        # Always print these on failure
        if not response.ok:
            print("STATUS:", response.status_code)
            print("HEADERS:", dict(response.headers))
            print("BODY:", response.text)  # <-- critical
        response.raise_for_status()
        result = response.json()
    except requests.exceptions.HTTPError as e:
        # Still has access to the response object
        r = e.response
        print("HTTPError:", e)
        if r is not None:
            print("STATUS:", r.status_code)
            print("HEADERS:", dict(r.headers))
            print("BODY:", r.text)
        raise
        
    response.raise_for_status()
    result = response.json()
    
    print("Job created:")
    print(result)
    
    out=result
    
    if "jobId" in result.keys():
    
        job_id = result["jobId"]
    
        # **************** Secondary request - GET
        
        #status_url = f"https://qt2gq77oik.execute-api.eu-central-1.amazonaws.com/prod/invoke/{job_id}"
        status_url = f"{api_url}/{job_id}"
        
        headers = {
            "x-api-key": api_key
        }
        
        while True:
            try:
                response = requests.get(status_url, headers=headers, timeout=5)
                if not response.ok:
                    print("STATUS:", response.status_code)
                    print("HEADERS:", dict(response.headers))
                    print("BODY:", response.text)  # <-- critical
                response.raise_for_status()
                body = response.json()
            
                print(body)
            
                if body["status"] in ("SUCCEEDED", "FAILED"):
                    break
            except Exception as e:
                if show: print(f'GET request error: {e}')
                r = e.response
                if r is not None:
                    print("STATUS:", r.status_code)
                    print("HEADERS:", dict(r.headers))
                    print("BODY:", r.text)
                    
            time.sleep(2)
        
        # Print the response       
        status = body.get("status")
        
        if status == "SUCCEEDED":
            out=body.get("result")
        else:
            # bubble up the worker error message + trace
            out=f'API call unsuccessful with status {status}: {body.get("error", {}).get("message", "Job failed")}'
            raise RuntimeError(out)
    else:
        raise Exception(f'!!!ERROR: Response to intitial async API POST call does not include JobId key: {out}')
            
    return(out)

def call_bedrock(prompt=None,model='us.amazon.nova-lite-v1:0', api_url=None, api_key=None, system_prompt='You are a helpful assistant',
                max_tokens=1024, temp=0.1, tools_desc=None, web_grounding=0,credentials_path='..//Admin//credentials.txt',async_api=0,show=0):
    
    '''
    MODELS
    'us.meta.llama3-2-11b-instruct-v1:0'
    
    'eu.amazon.nova-lite-v1:0' - WORKS US & EU
    'eu.amazon.nova-pro-v1:0' - WORKS US & EU
    
    'us.amazon.nova-premier-v1:0' - WORKS US & EU
    'us.amazon.nova-2-lite-v1:0' - WORKS, US & EU
    
    'anthropic.claude-fable-5'
    'anthropic.claude-opus-4-8'
    'us.anthropic.claude-opus-4-7'
    'us.anthropic.claude-sonnet-4-6'
    'us.anthropic.claude-opus-4-6-v1'
    'us.anthropic.claude-haiku-4-5-v1:0'
    'us.anthropic.claude-sonnet-4-5-20250929-v1:0'
    'eu.anthropic.claude-sonnet-4-20250514-v1:0' - WORKS US & EU
    'eu.anthropic.claude-3-7-sonnet-20250219-v1:0' - WORKS  US & EU
    'eu.anthropic.claude-3-5-sonnet-20241022-v2:0' - WORKS US & EU
    
    'eu.anthropic.claude-3-5-sonnet-20240620-v1:0' - WORKS US & EU

    # ********** PARAMETERS
    temp - low -> deterministic, high -> creative
    max_tokens - <=0 triggers max output token mapping routine

    SYNC: 'https://gzukpgh4eg.execute-api.eu-central-1.amazonaws.com/prod/invoke'
    ASYNC: 'https://qt2gq77oik.execute-api.eu-central-1.amazonaws.com/prod/invoke'

    TOOLS_DESC - No tools: None, otherwise dictionary in standard Bedrock format, e.g.:
    tools_desc= {
            "tools": [
                {
                    "toolSpec": {
                        "name": "tavily_search_tool",
                        "description": "Searches Tavily for research topics.",
                        "inputSchema": {
                            "json": {
                                "type": "object",
                                "properties": {
                                    "query": {"type": "string"}
                                },
                                "required": ["query"]
                            }
                        }
                    }
                },
                {
                    "toolSpec": {
                        "name": "arxiv_search_tool",
                        "description": "Searches arXiv for academic papers.",
                        "inputSchema": {
                            "json": {
                                "type": "object",
                                "properties": {
                                    "query": {"type": "string"}
                                },
                                "required": ["query"]
                            }
                        }
                    }
                }
            ]
        }
    
    # Test
    a=call_bedrock_api(prompt=None,model='eu.amazon.nova-lite-v1:0', api_url=api_url, api_key=api_key, show=1)
    
    '''
    
    models_with_grounding=['nova-2-lite','nova-premier','nova-2-pro']
    
    # Check environment type - AWS Sagemaker vs local
    aws=0
    try:
        if "HOME" in os.environ:
            if "ec2-user" in os.environ.get("HOME"):
                aws=1
                if show: print('********** Running on AWS')
    except:
        if show: print('********** Running on Locally')

    if web_grounding:
        if any([m in model for m in models_with_grounding]):
            model=model.replace('eu.','us.')
            if show: print(f'********** Web grounding enabled for model {model}')
            if tools_desc:
                tools_desc["tools"].append({"systemTool": {"name": "nova_grounding"}})
            else:
                tools_desc={"tools": [
                                        {
                                            "systemTool": {"name": "nova_grounding"}
                                        }
                                    ]
                                }
        else:
            print('WARNING: Web grounding only available for Amazon NOVA Lite 2, Pro 2 and Premier models!!!')

    if show:
        print(f'PROMPT: {prompt}')
        print(f'TOOLS DESC: {tools_desc}')
        print(f'INFERENCE PROFILE / MODEL ID: {model}')
    
    # Call bedrock
    
    if aws:
        o,p=call_bedrock_aws(prompt=prompt,model_id=model, system_prompt=system_prompt,max_tokens=max_tokens, temp=temp, 
                             tools_desc=tools_desc, show=show)

    else:
        if async_api:
            if show: print('********** NOT Running on AWS - using ASYNC API')
            #AWS LAMBDA: arn:aws:lambda:eu-central-1:730335465043:function:aws-bedrock-api-v0-BedrockLambdaFunction-RjV60Y6J1th3
            #APP: bedrock-async-api-v0
            
            if (api_url is None) | (api_key is None):
                try:
                    #import dbConnect as dbc
                    odb_creds=dbc.jsonpass(pattern=None,fn=credentials_path)
                    api_url =odb_creds['bedrock_api_async']['API_BASE']
                    api_key =odb_creds['bedrock_api_async']['API_KEY']
                except Exception as e:
                    raise Exception(f'Cannot find suitable credentials: {e}')
            p=prompt
            o=call_bedrock_api_async(prompt=prompt,model=model, api_url=api_url, api_key=api_key, system_prompt=system_prompt,
                                max_tokens=max_tokens, temp=temp, tools_desc=tools_desc, show=show)
        else:
            if show: print('********** NOT Running on AWS - using API')
            #AWS LAMBDA: arn:aws:lambda:eu-central-1:730335465043:function:aws-bedrock-api-v0-BedrockLambdaFunction-RjV60Y6J1th3
            #APP: aws-bedrock-api-v0
            
            if (api_url is None) | (api_key is None):
                try:
                    #import dbConnect as dbc
                    odb_creds=dbc.jsonpass(pattern=None,fn=credentials_path)
                    api_url =odb_creds['bedrock_api']['API_BASE']
                    api_key =odb_creds['bedrock_api']['API_KEY']
                except Exception as e:
                    raise Exception(f'Cannot find suitable credentials: {e}')
                    
            o,p=call_bedrock_api(prompt=prompt,model=model, api_url=api_url, api_key=api_key, system_prompt=system_prompt,
                                max_tokens=max_tokens, temp=temp, tools_desc=tools_desc, show=show)
    if show:
        print(f'*************** RESPONSE \n {o}')
        print(f'*************** PROMPT \n {p}')
        
    return(o,p)

def bedrock_agent(prompt=None,model='us.amazon.nova-lite-v1:0', api_url=None, api_key=None, system_prompt='You are a helpful assistant',
                     max_tokens=1024, temp=0.1, tools=None, tools_desc=None, async_api=0,show=0):

    '''
    AWS LAMBDA: arn:aws:lambda:eu-central-1:730335465043:function:aws-bedrock-api-v0-BedrockLambdaFunction-RjV60Y6J1th3
    APP: aws-bedrock-api-v0
    
    MODELS
    'us.meta.llama3-2-11b-instruct-v1:0'
    
    'eu.amazon.nova-lite-v1:0' - WORKS US & EU
    'eu.amazon.nova-pro-v1:0' - WORKS US & EU

    'us.amazon.nova-2-lite-v1:0' - WORKS, US ONLY
    
    'anthropic.claude-fable-5'
    'anthropic.claude-opus-4-8'
    'us.anthropic.claude-opus-4-7'
    'us.anthropic.claude-sonnet-4-6'
    'us.anthropic.claude-opus-4-6-v1'
    'us.anthropic.claude-haiku-4-5-v1:0'
    'us.anthropic.claude-sonnet-4-5-20250929-v1:0'
    'eu.anthropic.claude-sonnet-4-20250514-v1:0' - WORKS US & EU
    'eu.anthropic.claude-3-7-sonnet-20250219-v1:0' - WORKS  US & EU
    'eu.anthropic.claude-3-5-sonnet-20241022-v2:0' - WORKS US & EU
    
    'eu.anthropic.claude-3-5-sonnet-20240620-v1:0' - WORKS US & EU

    # ********** PARAMETERS
    temp - low -> deterministic, high -> creative
    max_tokens - <=0 triggers max output token mapping routine
    TOOLS - for standard tools: list of tool names; for custom tools: dict -> {'tool_name':tool_function_call}, e.g. {'average':np.mean}
    TOOLS_DESC - for standard tools: None; for custom tools: dictionary in standard Bedrock format, e.g.:
    tools_desc= {
            "tools": [
                {
                    "toolSpec": {
                        "name": "tavily_search_tool",
                        "description": "Searches Tavily for research topics.",
                        "inputSchema": {
                            "json": {
                                "type": "object",
                                "properties": {
                                    "query": {"type": "string"}
                                },
                                "required": ["query"]
                            }
                        }
                    }
                },
                {
                    "toolSpec": {
                        "name": "arxiv_search_tool",
                        "description": "Searches arXiv for academic papers.",
                        "inputSchema": {
                            "json": {
                                "type": "object",
                                "properties": {
                                    "query": {"type": "string"}
                                },
                                "required": ["query"]
                            }
                        }
                    }
                }
            ]
        }
    '''

    # TOOLS = list => auto-generation of dictionary based on tools pre-defined in llm_tools_jt.py
    # TOOLS = None => no tools use
    # TOOLS = dict - custom tools, TOOLS_DESC needs to be also provided
    if isinstance(tools,list):
        
        tools_dict = get_tools_schema()
        
        # get all tools
        if len(tools)==0: 
            tools={}
            for t in tools_dict['tools']:
                tools[t['toolSpec']['name']] = eval(f"{t['toolSpec']['name']}")
            tools_desc=tools_dict.copy()
            if show: print(tools)
        else:
            tools_desc={"tools": []}
            tools0=tools.copy()
            tools={}
            for t in tools_dict['tools']:
                if t['toolSpec']['name'] in tools0: 
                    tools_desc['tools'].append(t)
                    tools[t['toolSpec']['name']] = eval(f"{t['toolSpec']['name']}")
    elif isinstance(tools,type(None)):
        tools_desc=None

    if tools_desc is not None:
        tools_desc['toolChoice']={"auto": {}}
        if show:
            print(tools_desc)
            print(tools)
    
    o,p=call_bedrock(prompt=prompt,model=model, api_url=api_url, api_key=api_key, system_prompt=system_prompt,
                     max_tokens=max_tokens, temp=temp, show=show, tools_desc=tools_desc,async_api=async_api)

    messages=p+[o]

    if show: 
        print('***MESSAGES AFTER INITIAL RUN')
        print(messages)
    
    #output_content = response['output']['message']['content']
    tool_use =0
    for block in o['content']:
        
        # Check if the model wants to use a tool
        if 'toolUse' in block:
            tool_use =1
            tool_use_data = block['toolUse']
            tool_id = tool_use_data['toolUseId']
            tool_name = tool_use_data['name']
            tool_input = tool_use_data['input'] # This is already a Python dict, no need for json.loads()

            if show:
                print(f"**** Tool use info: \n {tool_use_data} \n")
            
            tool_result = str(tools[tool_name](**tool_input))
            
            if show:
                print(f"**** Model requested tool: \n {tool_name} \n")
                print(f"**** Tool input: \n {tool_input} \n")
                print(f"**** Tool result: \n {tool_result} \n")
                
            messages.append({
                            "role": "assistant",  # the model or tool result role
                            "content": [
                                {"text": tool_result}
                            ]
                        })

    if tool_use:
        o,p=call_bedrock(prompt=messages,model=model, api_url=api_url, api_key=api_key, system_prompt=system_prompt,
                     max_tokens=max_tokens, temp=temp, show=show, tools_desc=tools_desc,async_api=async_api)

        messages=messages+p+[o]
        
        if show: 
            print('***MESSAGES WITH TOOL USE')
            print(messages)
            
    return(messages)


def bedrock_agent_reflect(prompt=None,model='us.amazon.nova-lite-v1:0', api_url=None, api_key=None, system_prompt='You are a helpful assistant',
                         max_tokens=1024, temp=0.1, tools=None, tools_desc=None, reflect=None, async_api=0,show=0):

    '''
    AWS LAMBDA: arn:aws:lambda:eu-central-1:730335465043:function:aws-bedrock-api-v0-BedrockLambdaFunction-RjV60Y6J1th3
    APP: aws-bedrock-api-v0
    
    MODELS
    'us.meta.llama3-2-11b-instruct-v1:0'
    
    'eu.amazon.nova-lite-v1:0' - WORKS US & EU
    'eu.amazon.nova-pro-v1:0' - WORKS US & EU

    'us.amazon.nova-2-lite-v1:0' - WORKS, US ONLY
    
    'anthropic.claude-fable-5'
    'anthropic.claude-opus-4-8'
    'us.anthropic.claude-opus-4-7'
    'us.anthropic.claude-sonnet-4-6'
    'us.anthropic.claude-opus-4-6-v1'
    'us.anthropic.claude-haiku-4-5-v1:0'
    'us.anthropic.claude-sonnet-4-5-20250929-v1:0'
    'eu.anthropic.claude-sonnet-4-20250514-v1:0' - WORKS US & EU
    'eu.anthropic.claude-3-7-sonnet-20250219-v1:0' - WORKS  US & EU
    'eu.anthropic.claude-3-5-sonnet-20241022-v2:0' - WORKS US & EU
    
    'eu.anthropic.claude-3-5-sonnet-20240620-v1:0' - WORKS US & EU

    # ********** PARAMETERS
    temp - low -> deterministic, high -> creative
    max_tokens - <=0 triggers max output token mapping routine
    REFLECT: None for no reflection, string containing reflection model name (e.g. 'eu.anthropic.claude-sonnet-4-20250514-v1:0')
    TOOLS - for standard tools: list of tool names; for custom tools: dict -> {'tool_name':tool_function_call}, e.g. {'average':np.mean}
    TOOLS_DESC - for standard tools: None; for custom tools: dictionary in standard Bedrock format, e.g.:
    tools_desc= {
            "tools": [
                {
                    "toolSpec": {
                        "name": "tavily_search_tool",
                        "description": "Searches Tavily for research topics.",
                        "inputSchema": {
                            "json": {
                                "type": "object",
                                "properties": {
                                    "query": {"type": "string"}
                                },
                                "required": ["query"]
                            }
                        }
                    }
                },
                {
                    "toolSpec": {
                        "name": "arxiv_search_tool",
                        "description": "Searches arXiv for academic papers.",
                        "inputSchema": {
                            "json": {
                                "type": "object",
                                "properties": {
                                    "query": {"type": "string"}
                                },
                                "required": ["query"]
                            }
                        }
                    }
                }
            ]
        }
    '''
    
    messages=bedrock_agent(prompt=prompt,model=model, api_url=api_url, api_key=api_key, system_prompt=system_prompt,
                         max_tokens=max_tokens, temp=temp, show=show, tools=tools, tools_desc=tools_desc,async_api=async_api)

    if not isinstance(messages,list):messages=[messages]
    
    if show: print(f'MESSAGES - ORIGINAL \n {messages}')
        
    if reflect:

        prompt_reflect = f"""
        You are a skillful and knowledgeable reviewer and refiner.
        
        User asked:
        {prompt}
        
        Original answer:
        {messages}
        
        Step 1: Briefly evaluate if the original answer fully answers the user's question.
        Step 2: If improvement is needed, provide a refined answer. If the original answer is already correct, return it unchanged.
        
        Return STRICT JSON with two fields:
        {{
          "feedback": "<1-3 sentences explaining the gap or confirming correctness>",
          "refined_answer": "<final answer>"
        }}
        """
        
        message=bedrock_agent(prompt=prompt_reflect,model=reflect, api_url=api_url, api_key=api_key, system_prompt=system_prompt,
                               max_tokens=max_tokens, temp=temp, show=show, tools=tools, tools_desc=tools_desc,async_api=async_api)

        messages.append(message)
        
        if show: print(f'MESSAGES - REFINED \n {messages}')
        
    return(messages)

def call_ollama(query, model_code='gemma4:e4b', temp=0.95, sys_prompt=None,
                 show=0, max_tokens=10000, option='A', api_key=None):
    '''
    import os
    import openai

    option='A' -> Local Ollama (default)
        Local models e.g. ['gemma4:e4b', 'gemma4:e2b']
        Runs against your local Ollama server (http://localhost:11434/v1)

    option='B' -> Ollama Cloud
        Cloud-hosted models e.g. ['gpt-oss:120b-cloud', 'deepseek-v3.1:671b-cloud',
                                   'qwen3-coder:480b-cloud']
        Runs against Ollama's cloud API (https://ollama.com/v1)
        Requires an API key: set OLLAMA_API_KEY env var, or pass api_key= explicitly.
        Get a key at https://ollama.com/settings/keys (or run `ollama signin`).
    
    # Local (default, unchanged behavior)
    resp, text = call_ollama("Explain quicksort", model_code='gemma4:e4b', option='A')

    # Cloud — via env var
    # export OLLAMA_API_KEY="your_key_here"
    resp, text = call_ollama("Explain quicksort", model_code='gpt-oss:120b-cloud', option='B')

    # Cloud — passing key directly
    resp, text = call_ollama(
        "Explain quicksort",
        model_code='deepseek-v3.1:671b-cloud',
        option='B',
        api_key='your_key_here'
    )
    '''

    if option == 'B':
        # --- Ollama Cloud ---
        resolved_key = api_key or os.environ.get('OLLAMA_API_KEY')
        if not resolved_key:
            raise ValueError(
                "No Ollama Cloud API key found. Set the OLLAMA_API_KEY environment "
                "variable or pass api_key='...' explicitly."
            )
        client = openai.OpenAI(
            base_url='https://ollama.com/v1',
            api_key=resolved_key
        )
    elif option == 'A':
        # --- Local Ollama ---
        client = openai.OpenAI(
            base_url='http://localhost:11434/v1',
            api_key='ollama'  # required, but unused locally
        )
    else:
        raise ValueError("option must be 'A' (local) or 'B' (cloud)")

    if sys_prompt is None:
        sys_prompt = "You are a helpful assistant that answers questions asked by the user."

    prompts = [
        {"role": "system", "content": sys_prompt},
        {"role": "user", "content": query}
    ]

    if show: print(prompts)

    response = client.chat.completions.create(
        model=model_code,
        messages=prompts,
        temperature=temp,
        max_tokens=max_tokens
    )

    if show: print(response)

    final_response = response.choices[0].message.content

    if show: print(final_response)

    return response, final_response

def call_gpt(query, model_code='o4-mini', temp=0.95, sys_prompt=None, show=0,
             max_tokens=2000, web_search=0, odb_creds=None, reasoning=0):
    """OpenAI wrapper. Accepts either a plain string or a multimodal
    content list (text + image_url blocks) as `query`.

    Models: https://platform.openai.com/docs/pricing

    Available:
    gpt-4-0613
    gpt-4
    gpt-3.5-turbo
    gpt-4-1106-preview
    gpt-3.5-turbo-1106
    gpt-4-0125-preview
    gpt-4-turbo-preview
    gpt-3.5-turbo-0125
    gpt-4-turbo
    gpt-4-turbo-2024-04-09
    gpt-4o
    gpt-4o-2024-05-13
    gpt-4o-mini-2024-07-18
    gpt-4o-mini
    gpt-4o-2024-08-06
    chatgpt-4o-latest
    gpt-4o-2024-11-20
    gpt-4.5-preview
    gpt-4.5-preview-2025-02-27
    gpt-4.1-2025-04-14
    gpt-4.1
    gpt-4.1-mini-2025-04-14
    gpt-4.1-mini
    gpt-4.1-nano-2025-04-14
    gpt-4.1-nano
    gpt-3.5-turbo-16k    

    Cheapest (per token):
    - Nano models: GPT‑4.1 nano – $0.10 / $0.40
    - GPT‑4o mini – $0.15 / $0.60
    
    Mid-range:
    - GPT‑3.5 turbo – $3 / $6
    - GPT‑4.1 – $2 / $8
    - GPT‑4 turbo – $10 / $30
    
    Premium:
    - GPT‑4 – $30 / $60
    - GPT‑4o – $5 / $20 (value vs GPT‑4)
    
    Top‑end frontier:
    - GPT‑4.5 preview – $75 / $150
    
    https://platform.openai.com/docs/guides/deep-research
    https://platform.openai.com/docs/guides/tools-web-search?api-mode=responses&lang=python
    
    """

    if odb_creds is None:
        odb_creds = dbc.jsonpass(pattern=None, fn='..//Admin//credentials.txt')

    if 'openai' in odb_creds.keys():
        OPENAI_API_KEY = odb_creds['openai']['OPENAI_API_KEY']
    elif 'OPENAI_API_KEY' in odb_creds.keys():
        OPENAI_API_KEY = odb_creds['OPENAI_API_KEY']
    else:
        OPENAI_API_KEY = odb_creds['API_KEY']

    openai.api_key = OPENAI_API_KEY
    client = openai.OpenAI(api_key=OPENAI_API_KEY)

    if sys_prompt is None:
        sys_prompt = "You are a helpful assistant that answers questions asked by the user."

    # ---- Normalise query into a chat "content" value -------------------
    # If it's already a list of OpenAI content blocks, pass it straight through.
    is_multimodal = isinstance(query, list)

    user_content = query if is_multimodal else query

    prompts = [
        {"role": "system", "content": sys_prompt},
        {"role": "user", "content": user_content}
    ]

    if show:
        print(prompts)

    _is_reasoning_model = model_code.startswith('o') and model_code[1:2].isdigit()

    if web_search:
        # web_search path uses the Responses API which takes plain input.
        # For multimodal, fall back to chat.completions below instead.
        if is_multimodal:
            response = client.chat.completions.create(
                model=model_code,
                messages=prompts,
                temperature=temp,
                max_tokens=max_tokens
            )
            final_response = response.choices[0].message.content
        else:
            response = client.responses.create(
                model=model_code,
                tools=[{"type": "web_search_preview"}],
                input=query
            )
            final_response = response.output_text

    elif _is_reasoning_model or reasoning:
        _effort = 'high' if reasoning else 'medium'
        response = client.chat.completions.create(
            model=model_code,
            messages=prompts,
            max_completion_tokens=max_tokens,
            reasoning_effort=_effort
        )
        final_response = response.choices[0].message.content

    else:
        response = client.chat.completions.create(
            model=model_code,
            messages=prompts,
            temperature=temp,
            max_tokens=max_tokens
        )
        final_response = response.choices[0].message.content

    if show:
        print(response)
        print(final_response)

    return response, final_response

def call_gemini(query, model_code='gemini-2.0-flash', temp=0.95, sys_prompt=None,
                show=0, max_tokens=2000, web_search=0, odb_creds=None, reasoning=0):
    """Google Gemini wrapper (new `google-genai` SDK). Accepts either a
    plain string or a multimodal content list as `query`.

    Google Gemini API wrapper — drop-in replacement for call_gpt()
    Docs: https://ai.google.dev/gemini-api/docs

    pip install google-genai

    from google import genai
    from google.genai import types
    import base64 as _b64

    Gemini Models (March 2026)

    Frontier / Flagship Reasoning:
    - gemini-3.1-pro-preview        # most intelligent, deepest reasoning, agentic & multimodal
    
    Best Price / Performance (Reasoning-Capable):
    - gemini-3-flash-preview        # best price/performance, fast reasoning
    - gemini-2.5-flash              # stable, adaptive thinking (widely used default)
    
    Speed & Cost-Optimized:
    - gemini-3.1-flash-lite-preview # ultra-low cost, high-volume agentic tasks
    - gemini-2.0-flash              # next-gen speed & efficiency (legacy default)
    - gemini-2.0-flash-lite         # lowest latency, cheapest stable option
    
    Legacy / Maintenance (not recommended for new builds):
    - gemini-1.5-pro
    - gemini-1.5-flash
    - gemini-1.5-flash-8b
    
    API Pricing (USD per 1M tokens, input / output)
    
    - gemini-3.1-pro-preview        – $2.00–4.00 / $12.00–18.00 (tiered by prompt size)
    - gemini-3-flash-preview        – $0.50 / $3.00
    - gemini-3.1-flash-lite-preview – $0.25 / $1.50
    - gemini-2.5-pro                – $1.25–2.50 / $10.00–15.00
    - gemini-2.5-flash              – $0.30 / $2.50
    - gemini-2.0-flash              – $0.10 / $0.40
    - gemini-2.0-flash-lite         – $0.08 / $0.30
    - gemini-1.5-flash              – $0.075 / $0.30
    - gemini-1.5-flash-8b           – $0.0375 / $0.15
    
    
    Practical Recommendations:
    - Default reasoning model: gemini-3-flash-preview
    - Highest intelligence / agents: gemini-3.1-pro-preview
    - High-volume, lowest cost: gemini-3.1-flash-lite-preview
    - Cost-stable production workloads: gemini-2.5-flash
    
    Web search: built-in grounding tool, no extra per-search fee (unlike OpenAI)
    https://ai.google.dev/gemini-api/docs/grounding
    """

    if odb_creds is None:
        odb_creds = dbc.jsonpass(pattern=None, fn='..//Admin//credentials.txt')

    if 'gemini' in odb_creds.keys():
        GEMINI_API_KEY = odb_creds['gemini']['GEMINI_API_KEY']
    elif 'GEMINI_API_KEY' in odb_creds.keys():
        GEMINI_API_KEY = odb_creds['GEMINI_API_KEY']
    else:
        GEMINI_API_KEY = odb_creds['API_KEY']

    client = genai.Client(api_key=GEMINI_API_KEY)

    if sys_prompt is None:
        sys_prompt = "You are a helpful assistant that answers questions asked by the user."

    # ---- Normalise query into a list of Gemini Parts -------------------
    def _to_parts(q):
        # Plain string -> single text part
        if isinstance(q, str):
            return [types.Part.from_text(text=q)]

        parts = []
        for item in q:
            if not isinstance(item, dict):
                parts.append(types.Part.from_text(text=str(item)))
                continue

            # Text block from pdf_scan_to_text: {"text": "..."}
            if "text" in item:
                parts.append(types.Part.from_text(text=item["text"]))

            # Inline image block: {"inline_data": {"mime_type":..,"data": b64}}
            elif "inline_data" in item:
                inline = item["inline_data"]
                raw = inline["data"]
                # data may be base64 string -> decode to bytes
                if isinstance(raw, str):
                    raw = _b64.b64decode(raw)
                parts.append(
                    types.Part.from_bytes(
                        data=raw,
                        mime_type=inline.get("mime_type", "image/png")
                    )
                )
        return parts

    contents = _to_parts(query)

    # ---- Build generation config --------------------------------------
    cfg_kwargs = dict(
        temperature=temp,
        max_output_tokens=max_tokens,
        system_instruction=sys_prompt,
    )

    if reasoning:
        # 2.5 uses thinking_budget; 3.x uses thinking_level.
        parts_of_name = model_code.split('-')
        if '3' in (parts_of_name[1:2] or ['']):
            cfg_kwargs['thinking_config'] = types.ThinkingConfig(thinking_level="HIGH")
        else:
            cfg_kwargs['thinking_config'] = types.ThinkingConfig(thinking_budget=10000)

    if web_search:
        cfg_kwargs['tools'] = [types.Tool(google_search=types.GoogleSearch())]

    config = types.GenerateContentConfig(**cfg_kwargs)

    if show:
        print(f"System: {sys_prompt}")
        print(f"Contents: {contents}")

    response = client.models.generate_content(
        model=model_code,
        contents=contents,
        config=config,
    )

    final_response = response.text

    if show:
        print(response)
        print(final_response)

    return response, final_response

def call_claude(query, model_code='claude-sonnet-4-6', temp=0.95, sys_prompt=None,
                show=0, max_tokens=2000, web_search=0, odb_creds=None, reasoning=0):
    """
    Anthropic Claude wrapper. Accepts either a plain string or a
    multimodal content list (text + image blocks) as `query`.
    
    Anthropic Claude API wrapper — drop-in replacement for call_gpt()
    Docs: https://docs.anthropic.com/en/docs

    pip install anthropic
    import anthropic

    Available models:
    claude-opus-4-6-v-1                  # most intelligent, complex reasoning
    claude-sonnet-4-6                # best price/performance balance (default)
    claude-haiku-4-5-20251001        # fastest, most cost-efficient

    Pricing (input / output per 1M tokens):
    claude-opus-4-6-v-1      – $15 / $75
    claude-sonnet-4-6    – $3  / $15
    claude-haiku-4-5     – $0.80 / $4

    Web search: built-in tool, charged per search + token fees
    https://docs.anthropic.com/en/docs/agents-and-tools/tool-use
    
    """

    if odb_creds is None:
        odb_creds = dbc.jsonpass(pattern=None, fn='..//Admin//credentials.txt')

    if 'anthropic' in odb_creds.keys():
        ANTHROPIC_API_KEY = odb_creds['anthropic']['ANTHROPIC_API_KEY']
    elif 'ANTHROPIC_API_KEY' in odb_creds.keys():
        ANTHROPIC_API_KEY = odb_creds['ANTHROPIC_API_KEY']
    else:
        ANTHROPIC_API_KEY = odb_creds['API_KEY']

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    if sys_prompt is None:
        sys_prompt = "You are a helpful assistant that answers questions asked by the user."

    # ---- Normalise query into Claude message content -------------------
    # Claude accepts content as a string OR a list of content blocks.
    # pdf_scan_to_text() already emits Anthropic-native blocks
    # ({"type":"text"...}, {"type":"image","source":{...}}), so pass through.
    messages = [{"role": "user", "content": query}]

    if show:
        print(f"System: {sys_prompt}")
        print(f"Messages: {messages}")

    if reasoning:
        _think_budget = min(max(1024, max_tokens // 2), max_tokens - 1)
        _think_max = max(max_tokens, _think_budget + 1024)

        kwargs = dict(
            model=model_code,
            max_tokens=_think_max,
            temperature=1,
            thinking={"type": "enabled", "budget_tokens": _think_budget},
            system=sys_prompt,
            messages=messages
        )
        if web_search:
            kwargs["tools"] = [{"type": "web_search_20250305", "name": "web_search"}]

        response = client.messages.create(**kwargs)

        final_response = ' '.join(
            block.text for block in response.content
            if hasattr(block, 'text') and getattr(block, 'type', '') == 'text'
        )

    elif web_search:
        response = client.messages.create(
            model=model_code,
            max_tokens=max_tokens,
            temperature=temp,
            system=sys_prompt,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=messages
        )
        final_response = ' '.join(
            block.text for block in response.content
            if hasattr(block, 'text')
        )

    else:
        response = client.messages.create(
            model=model_code,
            max_tokens=max_tokens,
            temperature=temp,
            system=sys_prompt,
            messages=messages
        )
        # First text block (skip any non-text blocks safely)
        final_response = next(
            (block.text for block in response.content if hasattr(block, 'text')),
            ''
        )

    if show:
        print(response)
        print(final_response)

    return response, final_response

def call_openrouter(query, model_code='openai/gpt-4o-mini', temp=0.95, sys_prompt=None,
                     show=0, max_tokens=2000, web_search=0, odb_creds=None, reasoning=0,
                     mode='id', free_model_pick='first'):
    
    """
    import requests
    import openai
    OpenRouter wrapper (OpenAI-compatible API). Accepts either a plain string or a
    multimodal content list (text + image_url blocks) as `query`.

    OpenRouter docs: https://openrouter.ai/docs
    Models & pricing: https://openrouter.ai/models

    Model selection is controlled by `mode`:

    mode='id'  (default)
        Use `model_code` exactly as supplied, e.g.:
            'openai/gpt-4o-mini'
            'anthropic/claude-3.5-sonnet'
            'meta-llama/llama-3.1-8b-instruct'
        Works for any paid or already-free model slug on OpenRouter.

    mode='free_tier'
        Ignores `model_code`. Queries OpenRouter's live /models endpoint,
        filters for models that are entirely free (prompt cost == 0 and
        completion cost == 0), and picks one automatically:
            free_model_pick='first'  -> first free model returned by the API
            free_model_pick='random' -> random free model (helps spread load
                                         across OpenRouter's free-tier rate limits)
            free_model_pick='<slug>' -> pick a specific slug out of the free list
                                         if you want to pin one manually

    mode='free'
        Takes `model_code` (a normal model slug/family, e.g.
        'meta-llama/llama-3.1-8b-instruct') and forces OpenRouter's
        zero-cost variant of it by appending ':free', i.e.
        'meta-llama/llama-3.1-8b-instruct:free'. Use this when you know
        the family you want but want the no-cost variant specifically.

    web_search=1
        Appends ':online' to the final resolved model slug, which enables
        OpenRouter's built-in web-search plugin for that request.

    reasoning=1
        Passed through as {"reasoning": {"effort": "high"}} via extra_body
        for models on OpenRouter that support reasoning effort control.

    ### Summary of changes

    1. **Renamed** `call_gpt` → `call_openrouter`, and switched the OpenAI client to point at OpenRouter's OpenAI-compatible base URL (`https://openrouter.ai/api/v1`) with an OpenRouter API key (pulled from `odb_creds['openrouter']['OPENROUTER_API_KEY']` — adjust the key path/name to match your `credentials.txt` structure).

    2. **Added a `mode` parameter** with three ways to pick a model:
    - `mode='id'` (default) — use `model_code` verbatim, e.g. `'openai/gpt-4o-mini'` or `'anthropic/claude-3.5-sonnet'`.
    - `mode='free_tier'` — calls OpenRouter's `/models` endpoint live, filters for models with `prompt` and `completion` pricing both `== 0`, and auto-picks one (`free_model_pick='first'|'random'|'<specific-slug>'`). This adapts automatically as OpenRouter's free-tier lineup changes.
    - `mode='free'` — takes a model family you specify (e.g. `'meta-llama/llama-3.1-8b-instruct'`) and forces the `:free` suffixed variant (`'meta-llama/llama-3.1-8b-instruct:free'`).

    3. **Web search**: since OpenRouter doesn't have the OpenAI Responses API's `web_search_preview` tool, I replaced that logic with OpenRouter's own mechanism — appending `:online` to the model slug, which enables their built-in search plugin.

    4. **Reasoning effort**: passed through via `extra_body={"reasoning": {"effort": "high"}}`, which is OpenRouter's convention for reasoning-capable models (o1/o3/QwQ/DeepSeek-R1 etc.), rather than OpenAI's `reasoning_effort` chat param.

    5. Dropped the OpenAI-specific `_is_reasoning_model` name-prefix check (`o1`, `o3`, `o4-mini`...) since that logic was OpenAI-specific naming; reasoning is now purely controlled by the `reasoning=1` flag, which works across any OpenRouter reasoning-capable model.

    6. Kept the same multimodal handling (string vs. content-block list) and return signature: `(response, final_response)`.

    **Usage examples:**

    ```python
    # 1. Specific model ID
    call_openrouter("Explain quantum tunnelling", mode='id', model_code='anthropic/claude-3.5-sonnet')

    # 2. Auto-pick any $0 model currently free on OpenRouter
    call_openrouter("Summarize this article...", mode='free_tier', free_model_pick='random')

    # 3. Force the :free variant of a specific family
    call_openrouter("Write a haiku", mode='free', model_code='meta-llama/llama-3.1-8b-instruct')
    ```

    **Note:** you'll need to add an `openrouter` block to your credentials file, e.g.:
    ```json
    {"openrouter": {"OPENROUTER_API_KEY": "sk-or-...", "SITE_URL": "https://yourapp.com", "APP_NAME": "YourApp"}}
    ```
    """

    if odb_creds is None:
        odb_creds = dbc.jsonpass(pattern=None, fn='..//Admin//credentials.txt')

    if 'openrouter' in odb_creds.keys():
        OPENROUTER_API_KEY = odb_creds['openrouter']['OPENROUTER_API_KEY']
    elif 'OPENROUTER_API_KEY' in odb_creds.keys():
        OPENROUTER_API_KEY = odb_creds['OPENROUTER_API_KEY']
    else:
        OPENROUTER_API_KEY = odb_creds['API_KEY']

    BASE_URL = "https://openrouter.ai/api/v1"

    client = openai.OpenAI(
        api_key=OPENROUTER_API_KEY,
        base_url=BASE_URL,
    )

    # Optional but recommended by OpenRouter for attribution/rate-limit tracking
    extra_headers = {
        "HTTP-Referer": odb_creds.get('openrouter', {}).get('SITE_URL', ''),
        "X-Title": odb_creds.get('openrouter', {}).get('APP_NAME', ''),
    }

    # ---- Resolve which model slug to actually call ---------------------
    if mode == 'id':
        resolved_model = model_code

    elif mode == 'free':
        resolved_model = model_code if model_code.endswith(':free') else f"{model_code}:free"

    elif mode == 'free_tier':
        resp = requests.get(f"{BASE_URL}/models", headers={
            "Authorization": f"Bearer {OPENROUTER_API_KEY}"
        })
        resp.raise_for_status()
        all_models = resp.json().get('data', [])

        def _is_free(m):
            pricing = m.get('pricing', {})
            try:
                return float(pricing.get('prompt', 1)) == 0.0 and float(pricing.get('completion', 1)) == 0.0
            except (TypeError, ValueError):
                return False

        free_models = [m['id'] for m in all_models if _is_free(m)]

        if not free_models:
            raise RuntimeError("No free-tier ($0) models currently available on OpenRouter.")

        if free_model_pick == 'first':
            resolved_model = free_models[0]
        elif free_model_pick == 'random':
            import random
            resolved_model = random.choice(free_models)
        elif free_model_pick in free_models:
            resolved_model = free_model_pick
        else:
            resolved_model = free_models[0]

    else:
        raise ValueError("mode must be one of: 'id', 'free_tier', 'free'")

    if web_search and not resolved_model.endswith(':online'):
        resolved_model = f"{resolved_model}:online"

    if sys_prompt is None:
        sys_prompt = "You are a helpful assistant that answers questions asked by the user."

    # ---- Normalise query into a chat "content" value -------------------
    is_multimodal = isinstance(query, list)
    user_content = query if is_multimodal else query

    prompts = [
        {"role": "system", "content": sys_prompt},
        {"role": "user", "content": user_content}
    ]

    if show:
        print(f"Resolved model: {resolved_model}")
        print(prompts)

    extra_body = {}
    if reasoning:
        extra_body["reasoning"] = {"effort": "high"}

    response = client.chat.completions.create(
        model=resolved_model,
        messages=prompts,
        temperature=temp,
        max_tokens=max_tokens,
        extra_headers=extra_headers,
        extra_body=extra_body if extra_body else None,
    )
    final_response = response.choices[0].message.content

    if show:
        print(response)
        print(final_response)

    return response, final_response

def genai_master(query,provider='aws',model_code=None,temp=0.95,max_tokens=4096,odb_creds=None,web_search=0,
              sys_prompt=None,tools_desc=None,reasoning=0,async_api=False,ollama_fallback=0,agentic=0,tools_import=[],
              poll_interval=2.0,poll_timeout=300,react=0,max_loops=10,new_session=0,show=0):

    '''
    genai_master(query, provider='aws', model_code=None, temp=0.95, max_tokens=1000,
                 odb_creds=None, show=0, web_search=0, sys_prompt=None, tools_desc=None,
                 reasoning=0)

    REASONING (0=off, 1=on):
        OpenAI:    reasoning_effort='high', uses max_completion_tokens (o-series models auto-detected)
        Anthropic: extended thinking with budget_tokens (temperature forced to 1)
        Gemini:    thinking_config with thinking_budget (2.5) or thinking_level (3.x)
        Bedrock:   not supported through API proxy (use Claude/Nova natively)
        Ollama:    not supported (some models like deepseek-r1 reason natively)

    # Standard call (no reasoning)
    r, a = genai_master("Explain quantum computing", provider='openai')

    # With reasoning enabled
    r, a = genai_master("Solve this complex math proof", provider='claude', reasoning=1)

    # Also works at the individual function level
    r, a = call_gpt("Explain P vs NP", reasoning=1)
    r, a = call_claude("Design a distributed system", reasoning=1)
    r, a = call_gemini("Analyze this algorithm", model_code='gemini-2.5-flash', reasoning=1)

    '''
    
    if show:
        print('''
              **********************************************
              GEN AI MASTER
              **********************************************
              ''')

    if odb_creds is None:
        odb_creds=dbc.jsonpass(pattern=None,fn='..//Admin//credentials.txt')
        
    r=None
    a=None

    if not isinstance(sys_prompt,str):
        sys_prompt='You are a helpful assistant'
    elif len(sys_prompt)<=25:
        sys_prompt='You are a helpful assistant'
    
    try:
    
        if any(term in provider.lower() for term in ("bedrock", "aws", "amazon","nova")):
            
            if isinstance(odb_creds,str):
                cred_pth=odb_creds # '..//Admin//credentials.txt' etc
                api_url=None
                api_key=None
            elif isinstance(odb_creds,dict):
                cred_pth=None
                if 'API_BASE' in odb_creds.keys():
                    api_url=odb_creds['API_BASE']
                    api_key=odb_creds['API_KEY']
                else:
                    if async_api:
                        api_url=odb_creds['bedrock_api_async']['API_BASE']
                        api_key=odb_creds['bedrock_api_async']['API_KEY']
                    else:
                        api_url=odb_creds['bedrock_api']['API_BASE']
                        api_key=odb_creds['bedrock_api']['API_KEY']
            if agentic:
                print('Calling Bedrock Agent...')
                r=call_bedrock_agent(query,model=model_code,credentials_path=cred_pth,api_url=api_url, api_key=api_key,temp=temp,max_tokens=max_tokens,
                       sys_prompt=sys_prompt, tools_import=tools_import,poll_interval=poll_interval,poll_timeout=poll_timeout,
                       react=react,max_loops=max_loops,new_session=new_session,show=show)
                r=r.message
                p=r
            else:
                print('Calling Bedrock LLM...')
                r,p=call_bedrock(prompt=query,model=model_code, api_url=api_url, api_key=api_key, system_prompt=sys_prompt,
                    max_tokens=max_tokens, temp=temp, tools_desc=tools_desc, credentials_path=cred_pth,async_api=async_api,
                    show=show,web_grounding=web_search)
           
            if show:
                print(r)
                print(p)

            if isinstance(r,str): r=json.loads(r)

            try:
                a=r['content'][0]['text']
            except:
                try:
                    # Claude reasoning is provided as element #1 so taking second one
                    a=r['content'][1]['text']
                except:
                    a=r
                
        elif any(term in provider.lower() for term in ('openai','gpt','chatgpt')):
            if agentic:
                print('Calling ChatGPT agent...')
                r,a=call_agent(query, provider, model = model_code, odb_creds=odb_creds,sys_prompt=sys_prompt,reasoning=reasoning,temp=temp,show=show)
            else:
                print('Calling ChatGPT service...')
                r,a=call_gpt(query, model_code=model_code,temp=temp,show=show,max_tokens=max_tokens,web_search=web_search,odb_creds=odb_creds,
                            sys_prompt=sys_prompt,reasoning=reasoning)
        elif any(term in provider.lower() for term in ('anthropic','claude','sonnet','opus','haiku')):
            if agentic:
                print('Calling Anthropic agent...')
                r,a=call_agent(query, provider, model = model_code, odb_creds=odb_creds,sys_prompt=sys_prompt,reasoning=reasoning,temp=temp,show=show)
            else:
                print('Calling Anthropic service...')
                r,a=call_claude(query, model_code=model_code, temp=temp, sys_prompt=sys_prompt, show=show, max_tokens=max_tokens, web_search=web_search,
                                odb_creds=odb_creds, reasoning=reasoning)
        elif any(term in provider.lower() for term in ('google','gemini','vertex','vertexai')):
            if agentic:
                print('Calling Google Gemini agent...')
                r,a=call_agent(query, provider, model = model_code, odb_creds=odb_creds,sys_prompt=sys_prompt,reasoning=reasoning,temp=temp,show=show)
            else:
                print('Calling Google Gemini service...')
                r,a=call_gemini(query, model_code=model_code, temp=temp, sys_prompt=sys_prompt, show=show, max_tokens=max_tokens, web_search=web_search,
                                odb_creds=odb_creds, reasoning=reasoning)
        elif any(term in provider.lower() for term in ('ollama','local','localhost')):
            if agentic:
                print('Calling local Ollama agent...')
                r,a=call_agent(query, provider, model = model_code, odb_creds=odb_creds,sys_prompt=sys_prompt,reasoning=reasoning,temp=temp,show=show)
            else:
                print('Calling local Ollama LLM...')
                r,a=call_ollama(query, model_code=model_code,temp=temp,sys_prompt=sys_prompt,show=show,max_tokens=max_tokens)
        elif any(term in provider.lower() for term in ('open-router','router','openrouter','open_router')):
            if agentic:
                print('Calling OpenRouter agent...')
                r,a=call_agent(query, provider, model = model_code, odb_creds=odb_creds,sys_prompt=sys_prompt,reasoning=reasoning,temp=temp,show=show)
            else:
                print('Calling OpenRouter LLM...')
                r,a=call_openrouter(query, model_code=model_code, temp=temp, sys_prompt=sys_prompt,show=show, max_tokens=max_tokens, 
                                    web_search=web_search, odb_creds=odb_creds, reasoning=reasoning,mode='id', free_model_pick='first')
        elif any(term in provider.lower() for term in ('hf','face','hugging')):
            if agentic:
                print('Calling Huggingface agent...')
                r,a=call_agent(query, provider, model = model_code, odb_creds=odb_creds,sys_prompt=sys_prompt,reasoning=reasoning,temp=temp,show=show)
            else:
                print('Calling Huggingface LLM...')
                #r,a=call_hf(query, model_code=model_code, temp=temp, sys_prompt=sys_prompt,show=show, max_tokens=max_tokens, odb_creds=odb_creds)
        elif any(term in provider.lower() for term in ('ollama-cloud','ollamacloud','ollama_cloud')):
            if agentic:
                print('Calling Ollama Cloud agent...')
                # FIXME
            else:
                print('Calling Ollama Cloud LLM...')
                # FIXME
        else:
            raise Exception('Invalid provider...defaulting to OLLAMA with default settings')
            
    except Exception as e:
        
        ee=str(e)+'\n\n'+str(traceback.format_exc())
        
        if ollama_fallback:
            if show:
                print(f'ERROR: {ee} - defaulting to vanilla OLLAMA...')
                exc_type, exc_value, exc_traceback = sys.exc_info()
                traceback.print_exception(exc_type, exc_value, exc_traceback,limit=5, file=sys.stdout)
            try:
                r,a=call_ollama(query)
            except:
                if show:
                    exc_type, exc_value, exc_traceback = sys.exc_info()
                    traceback.print_exception(exc_type, exc_value, exc_traceback,limit=5, file=sys.stdout)
        else:
            if show:
                print(f'ERROR: {ee}')
                exc_type, exc_value, exc_traceback = sys.exc_info()
                traceback.print_exception(exc_type, exc_value, exc_traceback,limit=5, file=sys.stdout)

    return(r,a)

class MaxToolCallsHook(HookProvider):
    """
    ReAct safety guard: stop execution if the agent attempts too many tool calls.

    Why: "max_loops" in classic ReAct implementations usually means a cap on
    Thought→Action→Observation cycles. In Strands, the cleanest proxy is
    "tool calls", because each Action typically corresponds to one tool call.
    Hooks are the supported extension mechanism in Strands. [3](https://arxiv.org/html/2603.05344v3)
    """

    def __init__(self, max_tool_calls: int):
        self.max_tool_calls = int(max_tool_calls)
        self._count = 0

    def register_hooks(self, registry: HookRegistry, **kwargs: Any) -> None:
        registry.add_callback(BeforeToolCallEvent, self._before_tool_call)

    def _before_tool_call(self, event: BeforeToolCallEvent) -> None:
        self._count += 1
        if self._count > self.max_tool_calls:
            # Raise to hard-stop the run (simple + explicit).
            raise RuntimeError(
                f"ReAct loop limit reached: tool_calls={self._count} > max_loops={self.max_tool_calls}. "
                "Increase max_loops or tighten prompt/tools to reduce looping."
            )

def define_strands_agent(model_id:str='eu.amazon.nova-2-lite-v1:0',agent_type:str='generic',
                         max_loops:int=10) -> Agent:
    
    """Build and return a Strands agent backed by AWS Bedrock.
    
    # ---------------------------------------------------------------------------
    # Example usage
    # ---------------------------------------------------------------------------

    # 1. Simple factual question (no tools needed)
    response = agent("What is the capital of France?")
    print(response)

    # 2. Date-aware question
    response = agent("What is today's date and what day of the week is it?")
    print(response)

    # 3. Math via the calculator
    response = agent("What is (145 * 37 + 892) / 12.5?")
    print(response)

    # 4. Python code execution for data analysis
    response = agent(
        "Using Python, generate a list of the first 20 Fibonacci numbers "
        "and compute their mean and standard deviation."
    )
    print(response)

    # 5. Web search for current information
    response = agent("What are the latest developments in quantum computing?")
    print(response)

    # 6. File operations
    response = agent(
        "Create a CSV file called sample_data.csv with 5 rows of fake "
        "employee data (name, department, salary) and then read it back "
        "and tell me the average salary."
    )
    print(response)

    # 7. Multi-step reasoning combining several tools
    response = agent(
        "Search the web for the current population of the 5 largest "
        "countries, then write a Python script that creates a bar chart "
        "of the results and saves it as population_chart.png."
    )
    print(response)

    TERMINATE / FREE RESOURCES: del agent
    
    """

    if model_id[0:2]=='eu':
        region_name='eu-central-1'
    else:
        region_name='us-east-1'
    
    boto3.setup_default_session(region_name=region_name)

    model = BedrockModel(
        model_id=model_id,
    )

    ALL_TOOLS=agt.agentic_toolsets(subset=agent_type)
    
    agent = Agent(
        model=model,
        tools=ALL_TOOLS,
        #max_tool_turns=max_loops,
        system_prompt=(
            "You are a helpful assistant with access to tools for running "
            "Python code, searching the web, reading/writing files, running "
            "shell commands, and more. Use the appropriate tool whenever it "
            "would help answer the user's question accurately. "
            "Be concise and direct."
        ),
        # Hooks are the supported way to enforce limits / policies. [3](https://arxiv.org/html/2603.05344v3)
        hooks=[MaxToolCallsHook(max_tool_calls=max_loops)],
    )
    
    return agent

def react_strands_agent(
    model_id: str = 'eu.amazon.nova-2-lite-v1:0',
    agent_type:str='generic',
    max_loops: int = 10,
) -> Agent:
    
    """
    Build and return a Strands agent backed by AWS Bedrock with ReAct behavior.

    ReAct (Reason + Act + Observe) is implemented in Strands by:
      - prompt guidance (Reason → Act → Observe loop),
      - sequential tool execution,
      - optional tool-call cap via hooks.

    NOTE: Current Strands Agents does not use `strands.agent.strategy.ReActStrategy`.
    Agent configuration is done via Agent(...) params like `hooks` and `tool_executor`. [1](https://strandsagents.com/latest/documentation/docs/api-reference/python/interrupt/)
    """

    # Region selection (keep your original logic)
    region_name = "eu-central-1" if model_id.startswith("eu.") or model_id.startswith("eu") else "us-east-1"
    boto3.setup_default_session(region_name=region_name)

    # Bedrock model (Strands provider)
    model = BedrockModel(model_id=model_id)

    # Your tool factory (unchanged)
    ALL_TOOLS=agt.agentic_toolsets(subset=agent_type)

    # ReAct-style system prompt
    react_system_prompt = (
        "You are a helpful assistant with access to tools for running Python code, "
        "searching the web, reading/writing files, running shell commands, and more.\n\n"
        "Use a ReAct loop to solve tasks:\n"
        "1) Reason: briefly decide what you need next.\n"
        "2) Act: if needed, call the single best tool for the next step.\n"
        "3) Observe: incorporate the tool result.\n"
        "Repeat until you can answer.\n\n"
        "Rules:\n"
        "- Only use tools when needed.\n"
        "- Prefer one tool call per step.\n"
        "- Stop as soon as you have enough information.\n"
        "- Be concise and direct in the final answer."
    )

    agent = Agent(
        model=model,
        tools=ALL_TOOLS,
        # ReAct tends to be step-wise; sequential executor makes order deterministic. [2](https://strands.readthedocs.io/en/latest/)
        tool_executor=SequentialToolExecutor(),
        # Hooks are the supported way to enforce limits / policies. [3](https://arxiv.org/html/2603.05344v3)
        hooks=[MaxToolCallsHook(max_tool_calls=max_loops)],
        system_prompt=react_system_prompt,
    )

    return agent

def call_bedrock_agent_aws(prompt,model='eu.amazon.nova-2-lite-v1:0',react=0,max_loops=10,new_session=0,
                       show=0):

    out=None

    # Invoke agent
    global agent
    
    if new_session: 
        agent=None
        
    if agent is None: 
        print('No active agent found - initiating...')

        if react:
            agent = react_strands_agent(model_id=model, max_loops=max_loops)
        else:
            agent = define_strands_agent(model_id=model, max_loops=max_loops)

    out=agent(prompt)

    if show: print(out)

    return(out)

class BedrockApiModel(Model):
    
    """
    Strands Model implementation that calls a remote async Bedrock proxy.

    The remote Lambda exposes:
        POST /invoke          -> {"jobId": ...}  (202)
        GET  /invoke/{jobId}  -> {"status": ..., "result": <converse output message>}

    Custom Strands Model provider that routes inference through a remote
    Bedrock async API (API Gateway + Lambda + SQS + DynamoDB).
    
    import json
    import time
    import logging
    from typing import Any, Iterable, Optional
    
    import requests
    
    try:
        from strands import Agent, tool
        from strands.types.models import Model
        from strands.types.content import Messages
        from strands.types.streaming import StreamEvent
        from strands.types.tools import ToolSpec
        from strands_tools import python_repl
    except:
        !pip install strands-agents, strands-agents-tools
        from strands import Agent, tool
        from strands.types.models import Model
        from strands.types.content import Messages
        from strands.types.streaming import StreamEvent
        from strands.types.tools import ToolSpec
        from strands_tools import python_repl
    
    logger = logging.getLogger(__name__)
    """

    def __init__(
        self,
        api_url: str,
        api_key: str,
        model_id: str = "eu.amazon.nova-lite-v1:0",
        *,
        region: Optional[str] = None,
        temperature: float = 0.1,
        max_tokens: int = 2048,
        poll_interval: float = 2.0,
        poll_timeout: float = 300.0,
        request_timeout: float = 15.0,
    ):
        self.api_url = api_url.rstrip("/")
        self.api_key = api_key
        self._config = {
            "model_id": model_id,
            "region": region or ("us-east-1" if model_id.startswith("us") else "eu-central-1"),
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        self.poll_interval = poll_interval
        self.poll_timeout = poll_timeout
        self.request_timeout = request_timeout

    # ---- Strands Model interface ------------------------------------------

    def get_config(self) -> dict:
        return self._config

    def update_config(self, **kwargs) -> None:
        self._config.update(kwargs)
    
    def format_request(self, messages, tool_specs=None, system_prompt=None) -> dict:

        def _sanitize(obj):
            """Recursively convert Decimal -> int/float so boto3 accepts the document."""
            from decimal import Decimal
            if isinstance(obj, Decimal):
                return int(obj) if obj == obj.to_integral_value() else float(obj)
            if isinstance(obj, dict):
                return {k: _sanitize(v) for k, v in obj.items()}
            if isinstance(obj, (list, tuple)):
                return [_sanitize(v) for v in obj]
            return obj

        payload: dict[str, Any] = {
            "prompt": _sanitize(messages),          # <-- KEY FIX: sanitize the messages
            "modelId": self._config["model_id"],
            "region": self._config["region"],
            "system_prompt": system_prompt or "You are a helpful assistant",
            "temp": self._config["temperature"],
            "max_tokens": self._config["max_tokens"],
        }

        if tool_specs:
            payload["tools_desc"] = _sanitize({
                "tools": [{"toolSpec": spec} for spec in tool_specs]
            })

        return payload

    def format_chunk(self, event: dict) -> StreamEvent:
        """Pass through already-normalized stream events."""
        return event

    # ---- Core call: synchronous remote, faked as a single 'stream' --------
    '''
    def stream(self, request: dict) -> Iterable[StreamEvent]:
        """
        Strands drives the agent loop via stream(). Since the remote API is
        request/response (not token-streaming), we run the job to completion
        and then emit a single, well-formed Converse stream sequence.
        """
        output_message = self._invoke_remote(request)
        yield from self._synthesize_stream(output_message)
    '''
    
    async def stream(
        self,
        messages: Messages,
        tool_specs: Optional[list[ToolSpec]] = None,
        system_prompt: Optional[str] = None,
        **kwargs: Any,
    ) -> AsyncIterable[StreamEvent]:
        """
        Strands drives the agent loop via stream(). Since the remote API is
        request/response (not token-streaming), we run the job to completion
        and then emit a single, well-formed Converse stream sequence.
        """
        # Build the payload the Lambda expects
        request = self.format_request(messages, tool_specs, system_prompt)
    
        # Run the (blocking) remote call without blocking the event loop
        output_message = await asyncio.to_thread(self._invoke_remote, request)
    
        for event in self._synthesize_stream(output_message):
            yield event
            
    # ---- Structured output (required abstract method) ---------------------
    '''
    def structured_output(
        self,
        output_model: Type[T],
        prompt: Messages,
        system_prompt: Optional[str] = None,
        **kwargs: Any,
    ) -> Iterable[dict[str, T]]:
        """
        Produce a structured (Pydantic) output from the model.

        Strategy:
          1. Inject a system instruction containing the JSON schema of
             output_model so the model returns valid JSON.
          2. Run the remote job to completion (no token streaming).
          3. Extract the text, parse the JSON, validate against output_model.
          4. Yield {"output": <validated model instance>} as Strands expects.
        """
        # Build a schema-aware system prompt
        schema = json.dumps(output_model.model_json_schema(), indent=2)
        schema_instruction = (
            f"{system_prompt or 'You are a helpful assistant.'}\n\n"
            "You MUST respond with ONLY a single valid JSON object that conforms "
            "to the following JSON schema. Do not include markdown fences, "
            "explanations, or any text outside the JSON object.\n\n"
            f"JSON Schema:\n{schema}"
        )

        request = self.format_request(
            messages=prompt,
            tool_specs=None,
            system_prompt=schema_instruction,
        )

        output_message = self._invoke_remote(request)

        # Collect text from the response
        text_parts = [
            block["text"]
            for block in output_message.get("content", []) or []
            if "text" in block
        ]
        text = "".join(text_parts).strip()

        # Parse JSON (with fallback extraction)
        data = self._parse_json(text)

        # Validate against the Pydantic model
        result = output_model.model_validate(data)

        # Strands expects a generator yielding {"output": <model>}
        yield {"output": result}
    '''

    async def structured_output(
        self,
        output_model: Type[T],
        prompt: Messages,
        system_prompt: Optional[str] = None,
        **kwargs: Any,
    ) -> AsyncIterable[dict[str, T]]:
        schema = json.dumps(output_model.model_json_schema(), indent=2)
        schema_instruction = (
            f"{system_prompt or 'You are a helpful assistant.'}\n\n"
            "You MUST respond with ONLY a single valid JSON object that conforms "
            "to the following JSON schema. Do not include markdown fences, "
            "explanations, or any text outside the JSON object.\n\n"
            f"JSON Schema:\n{schema}"
        )
    
        request = self.format_request(
            messages=prompt,
            tool_specs=None,
            system_prompt=schema_instruction,
        )
    
        output_message = await asyncio.to_thread(self._invoke_remote, request)
    
        text_parts = [
            block["text"]
            for block in output_message.get("content", []) or []
            if "text" in block
        ]
        text = "".join(text_parts).strip()
        data = self._parse_json(text)
        result = output_model.model_validate(data)
    
        yield {"output": result}
    
    @staticmethod
    def _parse_json(text: str) -> dict:
        """Parse JSON from raw text, tolerating markdown fences / surrounding text."""
        # Strip common markdown code fences
        fenced = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, re.DOTALL)
        if fenced:
            text = fenced.group(1)

        try:
            return json.loads(text)
        except json.JSONDecodeError:
            # Fallback: grab the first {...} block
            match = re.search(r"\{.*\}", text, re.DOTALL)
            if not match:
                raise ValueError(f"Could not parse structured output from: {text!r}")
            return json.loads(match.group(0))
            
    # ---- Remote invocation -------------------------------------------------

    def _invoke_remote(self, payload: dict) -> dict:
        headers = {"Content-Type": "application/json", "x-api-key": self.api_key}

        # Safety net: guarantee no Decimal / non-JSON types remain
        clean_payload = make_json_safe(payload)   # reuse your existing utility

        resp = requests.post(
            self.api_url,
            headers=headers,
            data=json.dumps(clean_payload),        # <-- use cleaned payload
            timeout=self.request_timeout,
        )
        if not resp.ok:
            logger.error("POST failed %s: %s", resp.status_code, resp.text)
        resp.raise_for_status()
        result = resp.json()

        # Sync API path: result is already the converse output message
        if "jobId" not in result:
            return self._extract_output_message(result)

        # Async path: poll until terminal
        job_id = result["jobId"]
        status_url = f"{self.api_url}/{job_id}"
        deadline = time.monotonic() + self.poll_timeout

        while True:
            if time.monotonic() > deadline:
                raise TimeoutError(f"Job {job_id} timed out after {self.poll_timeout}s")

            r = requests.get(status_url, headers={"x-api-key": self.api_key},
                             timeout=self.request_timeout)
            r.raise_for_status()
            body = r.json()
            status = body.get("status")

            if status == "SUCCEEDED":
                return self._extract_output_message(body.get("result"))
            if status == "FAILED":
                err = body.get("error", {})
                raise RuntimeError(f"Bedrock job failed: {err.get('message', err)}")

            time.sleep(self.poll_interval)

    @staticmethod
    def _extract_output_message(result: Any) -> dict:
        """
        Normalize whatever the API returns into a Converse output 'message'
        dict: {"role": "assistant", "content": [...]}.
        """
        if not isinstance(result, dict):
            return {"role": "assistant", "content": [{"text": str(result)}]}

        # Lambda stores response.output.message directly
        if "role" in result and "content" in result:
            return result
        if "output" in result and "message" in result["output"]:
            return result["output"]["message"]
        if "message" in result:
            return result["message"]
        return {"role": "assistant", "content": [{"text": json.dumps(result)}]}

    # ---- Stream synthesis (Converse event protocol) -----------------------

    def _synthesize_stream(self, message: dict) -> Iterable[StreamEvent]:
        """
        Convert a complete output message into the event sequence Strands
        expects, correctly signalling tool_use vs end_turn stop reasons.
        """
        content_blocks = message.get("content", []) or []
        has_tool_use = any("toolUse" in b for b in content_blocks)

        yield {"messageStart": {"role": message.get("role", "assistant")}}

        for block in content_blocks:
            if "text" in block:
                yield {"contentBlockStart": {"start": {}}}
                yield {"contentBlockDelta": {"delta": {"text": block["text"]}}}
                yield {"contentBlockStop": {}}

            elif "toolUse" in block:
                tu = block["toolUse"]
                yield {"contentBlockStart": {
                    "start": {"toolUse": {
                        "toolUseId": tu["toolUseId"],
                        "name": tu["name"],
                    }}
                }}
                yield {"contentBlockDelta": {"delta": {
                    "toolUse": {"input": json.dumps(tu.get("input", {}))}
                }}}
                yield {"contentBlockStop": {}}

        yield {"messageStop": {
            "stopReason": "tool_use" if has_tool_use else "end_turn"
        }}

# --- credentials (reuse your dbConnect pattern if preferred) ---------------
def load_credentials(fn="..//Admin//credentials.txt",top_k="bedrock_api_async",url_k="API_BASE",key_k="API_KEY"):
    try:
        import dbConnect as dbc
        creds = dbc.jsonpass(pattern=None, fn=fn)
        a = creds[top_k]
    except:
        return None,None
    return a[url_k], a[key_k]

def call_bedrock_agent_api_async(user_prompt,model_id='eu.anthropic.claude-sonnet-4-6',temp=0.1,max_tokens=4096,poll_interval=2.0,
                                poll_timeout=300,sys_prompt="You are a helpful research assistant. Use tools when needed.",
                                tools_import=[],api_url=None, api_key=None,show=0):

    '''
    from strands import Agent
    from strands_tools import calculator, current_time
    from bedrock_api_model import BedrockApiModel
    from strands_tools import python_repl,file_read,file_write,editor,current_time,http_request,memory,agent_core_memory,think,batch,use_llm,calculator, diagram,handoff_to_user,retrieve,rss
    '''
    if (api_url is None) | (api_key is None):
        api_url, api_key = load_credentials()
    
    # --- build the custom model ------------------------------------------------
    model = BedrockApiModel(
        api_url=api_url,
        api_key=api_key,
        model_id=model_id,  # tool-capable
        temperature=temp,
        max_tokens=max_tokens,
        poll_interval=poll_interval,
        poll_timeout=poll_timeout,
    )
    
    # --- create the agent with tools -------------------------------------------
    # Check environment type - AWS Sagemaker vs local
    aws=0
    try:
        if "HOME" in os.environ:
            if "ec2-user" in os.environ.get("HOME"):
                aws=1
                if show: print('********** Running on AWS')
    except:
        if show: print('********** Running on Locally')
            
    # if on AWS:
    if aws:
        from strands_tools import python_repl,file_read,file_write,current_time,http_request, diagram
        recommended_tools=[python_repl,file_read,file_write,current_time,http_request, diagram]
    else:
        get_current_date,word_count,calculate,run_python_code,web_search,read_file,write_file,run_shell_command,json_query = agt.agentic_toolsets(subset='generic')
        recommended_tools=[get_current_date,word_count,calculate,run_python_code,web_search,read_file,write_file,json_query]
        
    tool_list = tools_import+recommended_tools
    
    agent = Agent(
        model=model,
        tools=tool_list,
        system_prompt=sys_prompt,
    )
    
    # --- run -------------------------------------------------------------------
    response = agent(user_prompt)
    
    if show:
        print("\n=== FINAL ANSWER ===")
        print(response)
    
    return(response)

def call_bedrock_agent(prompt,model='eu.amazon.nova-2-lite-v1:0',react=0,max_loops=10,new_session=0,
                       credentials_path='..//Admin//credentials.txt',temp=0.1,max_tokens=4096,
                       sys_prompt="You are a helpful research assistant. Use tools when needed.",
                       tools_import=[],poll_interval=2.0,poll_timeout=300,api_url=None, api_key=None,show=0):
    
    # Check environment type - AWS Sagemaker vs local
    aws=0
    try:
        if "HOME" in os.environ:
            if "ec2-user" in os.environ.get("HOME"):
                aws=1
                if show: print('********** Running on AWS')
    except:
        if show: print('********** Running on Locally')

    if show:
        pass
    
    # Call bedrock
    if 'agentcore' in model.lower():
        if aws:
            if show: print('********** Running on AWS - using existing agentcore agent instance')
            pass
        else:
            if show: print('********** NOT Running on AWS - using endpoint of an existing agentcore agent instance')
            pass
    else:
        if aws:
            if show: print('********** Running on AWS - invoking Strands agent directly')
            # def call_bedrock_agent_aws(prompt,model='eu.amazon.nova-2-lite-v1:0',react=0,max_loops=10,new_session=0,show=0)
            out=call_bedrock_agent_aws(prompt,model=model,react=react,max_loops=max_loops,new_session=new_session,show=show)

        else:
            if show: print('********** NOT Running on AWS - using async Bedrock API')
            #AWS LAMBDA: arn:aws:lambda:eu-central-1:730335465043:function:aws-bedrock-api-v0-BedrockLambdaFunction-RjV60Y6J1th3
            #APP: aws-bedrock-api-v0

            
            if (api_url is None) | (api_key is None):
                try:
                    import dbConnect as dbc
                    odb_creds=dbc.jsonpass(pattern=None,fn=credentials_path)
                    api_url =odb_creds['bedrock_api_async']['API_BASE']
                    api_key =odb_creds['bedrock_api_async']['API_KEY']
                except Exception as e:
                    raise Exception(f'Cannot find suitable credentials: {e}')

            # call_bedrock_agent_api_async(user_prompt,model_id='eu.anthropic.claude-sonnet-4-6',temp=0.1,max_tokens=4096,poll_interval=2.0,poll_timeout=300,sys_prompt="You are a helpful research assistant. Use tools when needed.",tools_import=[],show=0)
            out=call_bedrock_agent_api_async(prompt,model_id=model,temp=temp,max_tokens=max_tokens,poll_interval=poll_interval,poll_timeout=poll_timeout,
                                             sys_prompt=sys_prompt,tools_import=tools_import,api_url=api_url, api_key=api_key,show=show)

    if show:
        print(f'*************** RESPONSE \n {out}')
        print(f'*************** PROMPT \n {prompt}')
    
    return(out)

def define_ollama_agent(
    model_code: str = "gemma4:e4b",
    temp: float = 0.95,
    sys_prompt: Optional[str] = None,
    show: int = 0,
    max_tokens: int = 10000,
    max_loops: int = 10,
):
    """
    Create and return a Strands Agent backed by a local Ollama model.

    Notes
    -----
    - Requires Ollama server running (default host: http://localhost:11434). [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)
    - Strands Ollama provider is available via: pip install 'strands-agents[ollama]' strands-agents-tools. [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)

    Parameters
    ----------
    model_code : str
        Ollama model tag (e.g., 'gemma4:e4b', 'gemma4:e2b').
    temp : float
        Temperature for generation (higher = more random). [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)[2](https://strandsagents.com/docs/api/python/strands.models.ollama/)
    sys_prompt : Optional[str]
        System prompt for the agent.
    show : int
        Verbosity flag. Strands prints model output to stdout by default in basic usage. [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)
        (This function keeps `show` for API compatibility; see comment below.)
    max_tokens : int
        Maximum number of tokens to generate. [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)[2](https://strandsagents.com/docs/api/python/strands.models.ollama/)
    """

    # --- Imports here so module can be imported even if optional deps are missing ---
    try:
        from strands import Agent
        from strands.models.ollama import OllamaModel
    except Exception as e:
        raise ImportError(
            "Missing Strands Ollama dependencies. Install with:\n"
            "  pip install 'strands-agents[ollama]' strands-agents-tools\n"
            "and ensure Ollama is installed and running."
        ) from e

    if sys_prompt is None:
        sys_prompt = "You are a helpful assistant that answers questions asked by the user."

    # Basic parameter hygiene (safe clamps)
    try:
        temp = float(temp)
    except Exception:
        temp = 0.95
    temp = max(0.0, min(2.0, temp))

    try:
        max_tokens = int(max_tokens)
    except Exception:
        max_tokens = 2048
    max_tokens = max(1, min(200000, max_tokens))

    # Your tool factory (unchanged)
    ALL_TOOLS = define_strands_tools()

    # Strands OllamaModel expects the Ollama server host and the model identifier. [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)[2](https://strandsagents.com/docs/api/python/strands.models.ollama/)
    ollama_model = OllamaModel(
        host="http://localhost:11434",
        model_id=model_code,
        temperature=temp,
        max_tokens=max_tokens,
        # keep_alive="5m",            # optional; Strands documents this config option [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)
        # stop_sequences=["###"],     # optional; Strands documents this config option [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)[2](https://strandsagents.com/docs/api/python/strands.models.ollama/)
        # options={"top_k": 40},      # optional; Strands documents this config option [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)[2](https://strandsagents.com/docs/api/python/strands.models.ollama/)
    )

    # Create the Strands Agent with the configured model. [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)
    agent = Agent(
        model=ollama_model,
        system_prompt=sys_prompt,
        tools=ALL_TOOLS,  # optional: pass Strands tools here if needed [3](https://strandsagents.com/docs/user-guide/concepts/tools/custom-tools/)[4](https://github.com/strands-agents/sdk-python/blob/main/src/strands/tools/decorator.py)
        hooks=[MaxToolCallsHook(max_tool_calls=max_loops)],
    )

    # NOTE on `show`:
    # Strands basic usage prints to stdout by default. [1](https://strandsagents.com/docs/user-guide/concepts/model-providers/ollama/)
    # If you want show=0 to be truly silent in your environment, the typical pattern
    # is to configure the agent's callback/handler to a no-op in your app wiring.
    # (Keeping this function minimal and aligned with official OllamaModel usage.)

    return agent

def define_react_ollama_agent(
    model_id: str = 'llama3.1', # Common Ollama default
    max_loops: int = 10,
    base_url: str = "http://localhost:11434",
    max_tokens:int = 10000,
    temp: float = 0.95
) -> Agent:
    """
    Build and return a Strands agent backed by Ollama with ReAct behavior.
    """

    # 1. Initialize the Ollama model
    # Note: Ensure your local Ollama server is running.
    model = OllamaModel(model_id=model_id, base_url=base_url)

    # 2. Define tools (Placeholder for your define_strands_tools function)
    ALL_TOOLS = define_strands_tools()

    # 3. ReAct-style system prompt
    react_system_prompt = (
        "You are a helpful assistant with access to tools for running Python code, "
        "searching the web, reading/writing files, running shell commands, and more.\n\n"
        "Use a ReAct loop to solve tasks:\n"
        "1) Reason: briefly decide what you need next.\n"
        "2) Act: if needed, call the single best tool for the next step.\n"
        "3) Observe: incorporate the tool result.\n"
        "Repeat until you can answer.\n\n"
        "Rules:\n"
        "- Only use tools when needed.\n"
        "- Prefer one tool call per step.\n"
        "- Stop as soon as you have enough information.\n"
        "- Be concise and direct in the final answer."
    )

    # 4. Construct the Agent
    agent = Agent(
        model=model,
        tools=ALL_TOOLS,
        tool_executor=SequentialToolExecutor(),
        hooks=[MaxToolCallsHook(max_tool_calls=max_loops)],
        system_prompt=react_system_prompt,
        temperature=temp,
        max_tokens=max_tokens
    )

    return agent

def call_ollama_agent(prompt, model:str="gemma4:e4b",react:int=0,temp: float = 0.95,sys_prompt: Optional[str] = None,
                    show: int = 0,max_tokens: int = 10000,max_loops: int = 10,new_session:int=0):
    
    out=None

    # Invoke agent
    global agent
    
    if new_session: 
        agent=None
        
    if agent is None: 
        print('No active agent found - initiating...')

        if react:
            agent=define_react_ollama_agent(model_id=model,max_loops=max_loops,temp=temp,max_tokens=max_tokens,
                                        show=show,)
        else:
            agent=define_ollama_agent(model_code=model , temp=temp,sys_prompt=sys_prompt,max_tokens=max_tokens,
                                    max_loops=max_loops,show=show)

    out=agent(prompt)

    if show: print(out)

    return(out)

# ============================================================================
#  Multi-provider Strands agent builders / runners
#  Providers: Anthropic (Claude), Google (Gemini), OpenAI, OpenRouter
#
#  pip install 'strands-agents[anthropic,openai,litellm]' strands-agents-tools
#
# EXPECTED CREDENTIALS
# {
#  "anthropic":  {"ANTHROPIC_API_KEY": "sk-ant-..."},
#  "openai":     {"OPENAI_API_KEY": "sk-..."},
#  "gemini":     {"GEMINI_API_KEY": "AIza..."},
#  "openrouter": {"OPENROUTER_API_KEY": "sk-or-...", "SITE_URL": "...", "APP_NAME": "..."}
# ============================================================================

# ----------------------------------------------------------------------------
#  Shared: session registry (mirrors the global `agent` used by call_ollama_agent)
# ----------------------------------------------------------------------------
AGENT_SESSIONS: dict = {}

REACT_SYSTEM_PROMPT = (
    "You are a helpful assistant with access to tools for running Python code, "
    "searching the web, reading/writing files, and more.\n\n"
    "Use a ReAct loop to solve tasks:\n"
    "1) Reason: briefly decide what you need next.\n"
    "2) Act: if needed, call the single best tool for the next step.\n"
    "3) Observe: incorporate the tool result.\n"
    "Repeat until you can answer.\n\n"
    "Rules:\n"
    "- Only use tools when needed.\n"
    "- Prefer one tool call per step.\n"
    "- Stop as soon as you have enough information.\n"
    "- Be concise and direct in the final answer."
)

DEFAULT_SYSTEM_PROMPT = (
    "You are a helpful research assistant. Use tools when needed."
)


# ----------------------------------------------------------------------------
#  Shared helpers
# ----------------------------------------------------------------------------
def _load_creds(odb_creds=None, fn='..//Admin//credentials.txt'):
    """Lazy credential load, identical convention to call_gpt()/call_claude()."""
    try:
        if odb_creds is None:
            import dbConnect as dbc
            odb_creds = dbc.jsonpass(pattern=None, fn=fn)
    except:
        return None
    return odb_creds


def resolve_agent_tools(tools_import=None, show: int = 0, persona:str = 'general') -> List[Any]:
    """
    Tool-selection logic lifted from call_bedrock_agent_api_async() so every
    provider gets the same toolset.

    On AWS SageMaker -> heavier strands_tools (python_repl, diagram, ...).
    Locally         -> your sandboxed define_strands_tools() set.
    """
    tools_import = list(tools_import or [])

    aws = 0
    try:
        if "ec2-user" in os.environ.get("HOME", ""):
            aws = 1
    except Exception:
        aws = 0

    if show:
        print('********** Running on AWS' if aws else '********** Running locally')

    if aws:
        from strands_tools import (python_repl, file_read, file_write,
                                   current_time, http_request, diagram)
        recommended_tools = [python_repl, file_read, file_write,
                             current_time, http_request, diagram]
    else:
        recommended_tools = agt.agentic_toolsets(persona)
        if not isinstance(recommended_tools, (tuple, list)):
            recommended_tools = list(recommended_tools)

    return tools_import + recommended_tools


def _clamp(temp, max_tokens):
    """Parameter hygiene, as in define_ollama_agent()."""
    try:
        temp = float(temp)
    except Exception:
        temp = 0.95
    temp = max(0.0, min(2.0, temp))

    try:
        max_tokens = int(max_tokens)
    except Exception:
        max_tokens = 4096
    max_tokens = max(1, min(200000, max_tokens))

    return temp, max_tokens


def _pick_sys_prompt(sys_prompt, react):
    if react:
        return REACT_SYSTEM_PROMPT
    return sys_prompt or DEFAULT_SYSTEM_PROMPT


def _build_agent(model, sys_prompt, tool_list, max_loops, react=0, show=0):
    """Common Agent assembly + optional sequential ReAct executor."""
    from strands import Agent

    kwargs = dict(
        model=model,
        tools=tool_list,
        system_prompt=sys_prompt,
    )

    # MaxToolCallsHook / SequentialToolExecutor are optional in older setups
    try:
        kwargs["hooks"] = [MaxToolCallsHook(max_tool_calls=max_loops)]
    except Exception:
        if show:
            print("MaxToolCallsHook unavailable - running without loop cap.")

    if react:
        try:
            from strands.tools.executors import SequentialToolExecutor
            kwargs["tool_executor"] = SequentialToolExecutor()
        except Exception:
            try:
                kwargs["tool_executor"] = SequentialToolExecutor()  # module-level import
            except Exception:
                if show:
                    print("SequentialToolExecutor unavailable - using default executor.")

    return Agent(**kwargs)


# ============================================================================
#  1. ANTHROPIC  (Claude)
# ============================================================================
def define_claude_agent(
    model_code: str = 'claude-sonnet-5',
    temp: float = 0.95,
    sys_prompt: Optional[str] = None,
    show: int = 0,
    max_tokens: int = 4096,
    max_loops: int = 10,
    reasoning: int = 0,
    react: int = 0,
    tools_import: Optional[list] = None,
    odb_creds: Optional[dict] = None,
    extra_params: Optional[dict] = None,
):
    """
    Strands Agent backed by the Anthropic Messages API.

    Docs: https://strandsagents.com/docs/user-guide/concepts/model-providers/anthropic/
    Install: pip install 'strands-agents[anthropic]' strands-agents-tools

    reasoning=1 -> enables extended thinking. Anthropic requires temperature=1
                   when thinking is enabled, so temp is overridden.
    """
    try:
        from strands.models.anthropic import AnthropicModel
    except Exception as e:
        raise ImportError(
            "Missing Anthropic provider. Install with:\n"
            "  pip install 'strands-agents[anthropic]' strands-agents-tools"
        ) from e

    if odb_creds is None:
        odb_creds = dbc.jsonpass(pattern=None, fn='..//Admin//credentials.txt')
    elif isinstance(odb_creds, str): 
        odb_creds = _load_creds(odb_creds)
    
    if 'anthropic' in odb_creds.keys():
        ANTHROPIC_API_KEY = odb_creds['anthropic']['ANTHROPIC_API_KEY']
    elif 'ANTHROPIC_API_KEY' in odb_creds.keys():
        ANTHROPIC_API_KEY = odb_creds['ANTHROPIC_API_KEY']
    else:
        ANTHROPIC_API_KEY = odb_creds['API_KEY']

    temp, max_tokens = _clamp(temp, max_tokens)
    sys_prompt = _pick_sys_prompt(sys_prompt, react)

    params = {"temperature": temp}

    if reasoning:
        budget = min(max(1024, max_tokens // 2), max_tokens - 1)
        max_tokens = max(max_tokens, budget + 1024)
        params["temperature"] = 1                     # required with thinking
        params["thinking"] = {"type": "enabled", "budget_tokens": budget}

    if extra_params:
        params.update(extra_params)

    model = AnthropicModel(
        client_args={"api_key": ANTHROPIC_API_KEY},
        model_id=model_code,
        max_tokens=max_tokens,
        params=params,
    )

    tool_list = resolve_agent_tools(tools_import, show=show)

    if show:
        print(f"[claude-agent] model={model_code} params={params} tools={len(tool_list)}")

    return _build_agent(model, sys_prompt, tool_list, max_loops, react=react, show=show)


def call_claude_agent(
    prompt,
    model: str = 'claude-sonnet-5',
    temp: float = 0.95,
    sys_prompt: Optional[str] = None,
    show: int = 0,
    max_tokens: int = 4096,
    max_loops: int = 10,
    reasoning: int = 0,
    react: int = 0,
    tools_import: Optional[list] = None,
    odb_creds: Optional[dict] = None,
    new_session: int = 0,
):
    """Run (and cache) a Claude-backed Strands agent. Mirrors call_ollama_agent()."""
    key = 'claude'

    if new_session:
        AGENT_SESSIONS.pop(key, None)

    if AGENT_SESSIONS.get(key) is None:
        if show:
            print('No active Claude agent found - initiating...')
        AGENT_SESSIONS[key] = define_claude_agent(
            model_code=model, temp=temp, sys_prompt=sys_prompt, show=show,
            max_tokens=max_tokens, max_loops=max_loops, reasoning=reasoning,
            react=react, tools_import=tools_import, odb_creds=odb_creds,
        )

    out = AGENT_SESSIONS[key](prompt)

    if show:
        print("\n=== FINAL ANSWER ===")
        print(out)

    return out


# ============================================================================
#  2. GOOGLE  (Gemini)
# ============================================================================
def define_gemini_agent(
    model_code: str = 'gemini-3.6-flash',
    temp: float = 0.95,
    sys_prompt: Optional[str] = None,
    show: int = 0,
    max_tokens: int = 4096,
    max_loops: int = 10,
    reasoning: int = 0,
    react: int = 0,
    tools_import: Optional[list] = None,
    odb_creds: Optional[dict] = None,
    extra_params: Optional[dict] = None,
):
    """
    Strands Agent backed by Google Gemini.

    Prefers strands.models.gemini.GeminiModel (newer SDKs); falls back to
    LiteLLMModel with a 'gemini/<model>' slug on older installs.

    Install: pip install 'strands-agents[gemini]'   (or [litellm] for fallback)
    """

    if odb_creds is None:
        odb_creds = dbc.jsonpass(pattern=None, fn='..//Admin//credentials.txt')
    elif isinstance(odb_creds, str): 
        odb_creds = _load_creds(odb_creds)
    
    if 'gemini' in odb_creds.keys():
        GEMINI_API_KEY = odb_creds['gemini']['GEMINI_API_KEY']
    elif 'GEMINI_API_KEY' in odb_creds.keys():
        GEMINI_API_KEY = odb_creds['GEMINI_API_KEY']
    else:
        GEMINI_API_KEY = odb_creds['API_KEY']

    temp, max_tokens = _clamp(temp, max_tokens)
    sys_prompt = _pick_sys_prompt(sys_prompt, react)

    model = None

    # ---- Preferred: native Strands Gemini provider ------------------------
    try:
        from strands.models.gemini import GeminiModel

        params = {
            "temperature": temp,
            "max_output_tokens": max_tokens,
        }

        if reasoning:
            # 3.x uses thinking_level, 2.5 uses thinking_budget
            if model_code.split('-')[1:2] == ['3']:
                params["thinking_config"] = {"thinking_level": "HIGH"}
            else:
                params["thinking_config"] = {"thinking_budget": 10000}

        if extra_params:
            params.update(extra_params)

        model = GeminiModel(
            client_args={"api_key": GEMINI_API_KEY},
            model_id=model_code,
            params=params,
        )
        if show:
            print(f"[gemini-agent] native GeminiModel, params={params}")

    # ---- Fallback: LiteLLM route ------------------------------------------
    except Exception as e:
        if show:
            print(f"[gemini-agent] native provider unavailable ({e}); using LiteLLM.")
        try:
            from strands.models.litellm import LiteLLMModel
        except Exception as e2:
            raise ImportError(
                "Missing Gemini provider. Install one of:\n"
                "  pip install 'strands-agents[gemini]'\n"
                "  pip install 'strands-agents[litellm]'"
            ) from e2

        os.environ.setdefault("GEMINI_API_KEY", api_key)
        os.environ.setdefault("GOOGLE_API_KEY", api_key)

        params = {"temperature": temp, "max_tokens": max_tokens}
        if extra_params:
            params.update(extra_params)

        model = LiteLLMModel(
            client_args={"api_key": api_key},
            model_id=f"gemini/{model_code}",
            params=params,
        )

    tool_list = resolve_agent_tools(tools_import, show=show)

    if show:
        print(f"[gemini-agent] model={model_code} tools={len(tool_list)}")

    return _build_agent(model, sys_prompt, tool_list, max_loops, react=react, show=show)


def call_gemini_agent(
    prompt,
    model: str = 'gemini-3.6-flash',
    temp: float = 0.95,
    sys_prompt: Optional[str] = None,
    show: int = 0,
    max_tokens: int = 4096,
    max_loops: int = 10,
    reasoning: int = 0,
    react: int = 0,
    tools_import: Optional[list] = None,
    odb_creds: Optional[dict] = None,
    new_session: int = 0,
):
    """Run (and cache) a Gemini-backed Strands agent."""
    key = 'gemini'

    if new_session:
        AGENT_SESSIONS.pop(key, None)

    if AGENT_SESSIONS.get(key) is None:
        if show:
            print('No active Gemini agent found - initiating...')
        AGENT_SESSIONS[key] = define_gemini_agent(
            model_code=model, temp=temp, sys_prompt=sys_prompt, show=show,
            max_tokens=max_tokens, max_loops=max_loops, reasoning=reasoning,
            react=react, tools_import=tools_import, odb_creds=odb_creds,
        )

    out = AGENT_SESSIONS[key](prompt)

    if show:
        print("\n=== FINAL ANSWER ===")
        print(out)

    return out


# ============================================================================
#  3. OPENAI
# ============================================================================
def _is_openai_reasoning_model(model_code: str) -> bool:
    """o1 / o3 / o4-mini / gpt-5* families use reasoning params, not temperature."""
    m = model_code.lower()
    if m.startswith('o') and len(m) > 1 and m[1].isdigit():
        return True
    if m.startswith('gpt-5'):
        return True
    return False


def define_gpt_agent(
    model_code: str = 'gpt-4.1-mini',
    temp: float = 0.95,
    sys_prompt: Optional[str] = None,
    show: int = 0,
    max_tokens: int = 4096,
    max_loops: int = 10,
    reasoning: int = 0,
    react: int = 0,
    tools_import: Optional[list] = None,
    odb_creds: Optional[dict] = None,
    extra_params: Optional[dict] = None,
):
    """
    Strands Agent backed by the OpenAI Chat Completions API.

    Docs: https://strandsagents.com/docs/user-guide/concepts/model-providers/openai/
    Install: pip install 'strands-agents[openai]' strands-agents-tools

    Reasoning models (o3/o4-mini/gpt-5...) reject `temperature` and require
    `max_completion_tokens`, so those are swapped in automatically.
    """
    try:
        from strands.models.openai import OpenAIModel
    except Exception as e:
        raise ImportError(
            "Missing OpenAI provider. Install with:\n"
            "  pip install 'strands-agents[openai]' strands-agents-tools"
        ) from e
    
    if odb_creds is None: 
        odb_creds = dbc.jsonpass(pattern=None, fn='..//Admin//credentials.txt')
    elif isinstance(odb_creds, str): 
        odb_creds = _load_creds(odb_creds)
    
    if 'openai' in odb_creds.keys():
        OPENAI_API_KEY = odb_creds['openai']['OPENAI_API_KEY']
    elif 'OPENAI_API_KEY' in odb_creds.keys():
        OPENAI_API_KEY = odb_creds['OPENAI_API_KEY']
    else:
        OPENAI_API_KEY= odb_creds['API_KEY']

    temp, max_tokens = _clamp(temp, max_tokens)
    sys_prompt = _pick_sys_prompt(sys_prompt, react)

    if _is_openai_reasoning_model(model_code) or reasoning:
        params = {
            "max_completion_tokens": max_tokens,
            "reasoning_effort": "high" if reasoning else "medium",
        }
    else:
        params = {
            "temperature": temp,
            "max_tokens": max_tokens,
        }

    if extra_params:
        params.update(extra_params)

    model = OpenAIModel(
        client_args={"api_key": OPENAI_API_KEY},
        model_id=model_code,
        params=params,
    )

    tool_list = resolve_agent_tools(tools_import, show=show)

    if show:
        print(f"[gpt-agent] model={model_code} params={params} tools={len(tool_list)}")

    return _build_agent(model, sys_prompt, tool_list, max_loops, react=react, show=show)


def call_gpt_agent(
    prompt,
    model: str = 'gpt-4.1-mini',
    temp: float = 0.95,
    sys_prompt: Optional[str] = None,
    show: int = 0,
    max_tokens: int = 4096,
    max_loops: int = 10,
    reasoning: int = 0,
    react: int = 0,
    tools_import: Optional[list] = None,
    odb_creds: Optional[dict] = None,
    new_session: int = 0,
):
    """Run (and cache) an OpenAI-backed Strands agent."""
    key = 'gpt'

    if new_session:
        AGENT_SESSIONS.pop(key, None)

    if AGENT_SESSIONS.get(key) is None:
        if show:
            print('No active OpenAI agent found - initiating...')
        AGENT_SESSIONS[key] = define_gpt_agent(
            model_code=model, temp=temp, sys_prompt=sys_prompt, show=show,
            max_tokens=max_tokens, max_loops=max_loops, reasoning=reasoning,
            react=react, tools_import=tools_import, odb_creds=odb_creds,
        )

    out = AGENT_SESSIONS[key](prompt)

    if show:
        print("\n=== FINAL ANSWER ===")
        print(out)

    return out


# ============================================================================
#  4. OPENROUTER
# ============================================================================
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"


def resolve_openrouter_model(
    model_code: str,
    api_key: str,
    mode: str = 'id',
    free_model_pick: str = 'first',
    web_search: int = 0,
    tool_capable_only: int = 1,
) -> str:
    """
    Slug resolution extracted from call_openrouter() so agents can reuse it.

    mode='id'        -> use model_code verbatim
    mode='free'      -> force the ':free' variant of model_code
    mode='free_tier' -> query /models, filter $0 models, auto-pick

    tool_capable_only=1 additionally filters free models to those advertising
    'tools' support -- important for agents, since a model without function
    calling will silently never invoke your Strands tools.
    """
    if mode == 'id':
        resolved = model_code

    elif mode == 'free':
        resolved = model_code if model_code.endswith(':free') else f"{model_code}:free"

    elif mode == 'free_tier':
        resp = requests.get(
            f"{OPENROUTER_BASE_URL}/models",
            headers={"Authorization": f"Bearer {api_key}"},
            timeout=30,
        )
        resp.raise_for_status()
        all_models = resp.json().get('data', [])

        def _is_free(m):
            p = m.get('pricing', {})
            try:
                return float(p.get('prompt', 1)) == 0.0 and float(p.get('completion', 1)) == 0.0
            except (TypeError, ValueError):
                return False

        def _has_tools(m):
            params = m.get('supported_parameters') or []
            return ('tools' in params) or ('tool_choice' in params)

        free_models = [m['id'] for m in all_models if _is_free(m)]

        if tool_capable_only:
            tool_free = [m['id'] for m in all_models if _is_free(m) and _has_tools(m)]
            if tool_free:
                free_models = tool_free

        if not free_models:
            raise RuntimeError("No suitable free-tier ($0) models available on OpenRouter.")

        if free_model_pick == 'random':
            resolved = random.choice(free_models)
        elif free_model_pick in free_models:
            resolved = free_model_pick
        else:
            resolved = free_models[0]

    else:
        raise ValueError("mode must be one of: 'id', 'free_tier', 'free'")

    if web_search and not resolved.endswith(':online'):
        resolved = f"{resolved}:online"

    return resolved


def define_openrouter_agent(
    model_code: str = 'openai/gpt-4o-mini',
    temp: float = 0.95,
    sys_prompt: Optional[str] = None,
    show: int = 0,
    max_tokens: int = 4096,
    max_loops: int = 10,
    reasoning: int = 0,
    react: int = 0,
    web_search: int = 0,
    mode: str = 'id',
    free_model_pick: str = 'first',
    tools_import: Optional[list] = None,
    odb_creds: Optional[dict] = None,
    extra_params: Optional[dict] = None,
):
    """
    Strands Agent backed by OpenRouter via the OpenAI-compatible provider.

    Install: pip install 'strands-agents[openai]' strands-agents-tools

    Notes
    -----
    - Pick a *tool-capable* model; many free slugs lack function calling and
      will ignore your Strands tools entirely.
    - web_search=1 appends ':online' (OpenRouter's server-side search plugin),
      which composes fine with local Strands tools.
    - reasoning=1 is passed through as extra_body {"reasoning": {"effort": "high"}}.
    """
    try:
        from strands.models.openai import OpenAIModel
    except Exception as e:
        raise ImportError(
            "Missing OpenAI-compatible provider. Install with:\n"
            "  pip install 'strands-agents[openai]' strands-agents-tools"
        ) from e

    if odb_creds is None:
        odb_creds = dbc.jsonpass(pattern=None, fn='..//Admin//credentials.txt')
    elif isinstance(odb_creds, str): 
        odb_creds = _load_creds(odb_creds)
    
    if 'openrouter' in odb_creds.keys():
        OPENROUTER_API_KEY = odb_creds['openrouter']['OPENROUTER_API_KEY']
    elif 'OPENROUTER_API_KEY' in odb_creds.keys():
        OPENROUTER_API_KEY = odb_creds['OPENROUTER_API_KEY']
    else:
        OPENROUTER_API_KEY = odb_creds['API_KEY']

    temp, max_tokens = _clamp(temp, max_tokens)
    sys_prompt = _pick_sys_prompt(sys_prompt, react)

    resolved_model = resolve_openrouter_model(
        model_code=model_code,
        api_key=OPENROUTER_API_KEY,
        mode=mode,
        free_model_pick=free_model_pick,
        web_search=web_search,
    )

    params = {"temperature": temp, "max_tokens": max_tokens}
    if reasoning:
        params["extra_body"] = {"reasoning": {"effort": "high"}}
    if extra_params:
        params.update(extra_params)

    model = OpenAIModel(
        client_args={
            "api_key": OPENROUTER_API_KEY,
            "base_url": OPENROUTER_BASE_URL,
            "default_headers": {
                "HTTP-Referer": odb_creds.get('openrouter', {}).get('SITE_URL', ''),
                "X-Title": odb_creds.get('openrouter', {}).get('APP_NAME', ''),
            },
        },
        model_id=resolved_model,
        params=params,
    )

    tool_list = resolve_agent_tools(tools_import, show=show)

    if show:
        print(f"[openrouter-agent] resolved model={resolved_model} "
              f"params={params} tools={len(tool_list)}")

    return _build_agent(model, sys_prompt, tool_list, max_loops, react=react, show=show)

def call_openrouter_agent(
    prompt,
    model: str = 'openai/gpt-4o-mini',
    temp: float = 0.95,
    sys_prompt: Optional[str] = None,
    show: int = 0,
    max_tokens: int = 4096,
    max_loops: int = 10,
    reasoning: int = 0,
    react: int = 0,
    web_search: int = 0,
    mode: str = 'id',
    free_model_pick: str = 'first',
    tools_import: Optional[list] = None,
    odb_creds: Optional[dict] = None,
    new_session: int = 0,
):
    """Run (and cache) an OpenRouter-backed Strands agent."""
    key = 'openrouter'

    if new_session:
        AGENT_SESSIONS.pop(key, None)

    if AGENT_SESSIONS.get(key) is None:
        if show:
            print('No active OpenRouter agent found - initiating...')
        AGENT_SESSIONS[key] = define_openrouter_agent(
            model_code=model, temp=temp, sys_prompt=sys_prompt, show=show,
            max_tokens=max_tokens, max_loops=max_loops, reasoning=reasoning,
            react=react, web_search=web_search, mode=mode,
            free_model_pick=free_model_pick, tools_import=tools_import,
            odb_creds=odb_creds,
        )

    out = AGENT_SESSIONS[key](prompt)

    if show:
        print("\n=== FINAL ANSWER ===")
        print(out)

    return out


# ============================================================================
#  5. Unified dispatcher
# ============================================================================
AGENT_DEFAULT_MODELS = {
    'claude':     'claude-sonnet-5',
    'gemini':     'gemini-3.7-flash',
    'gpt':        'gpt-4.1-mini',
    'openrouter': 'openai/gpt-4o-mini',
    'ollama':     'gemma4:e4b',
    'bedrock':    'eu.anthropic.claude-sonnet-5',
}


def call_agent(prompt, provider: str = 'claude', model: Optional[str] = None, **kwargs):
    """
    One-line provider switch for tool-using agents.

    call_agent("Plot BTC vs ETH volatility", provider='gpt')
    call_agent("Research X and write a summary", provider='claude', reasoning=1)
    call_agent("Cheap agent run", provider='openrouter', mode='free_tier',
               free_model_pick='random')
    """
    provider = provider.lower()
    model = model or AGENT_DEFAULT_MODELS.get(provider)

    if provider in ('claude', 'anthropic'):
        return call_claude_agent(prompt, model=model, **kwargs)

    if provider in ('gemini', 'google'):
        return call_gemini_agent(prompt, model=model, **kwargs)

    if provider in ('gpt', 'openai'):
        return call_gpt_agent(prompt, model=model, **kwargs)

    if provider == 'openrouter':
        return call_openrouter_agent(prompt, model=model, **kwargs)

    if provider == 'ollama':
        return call_ollama_agent(prompt, model=model, **kwargs)

    if provider == 'bedrock':
        return call_bedrock_agent_api_async(prompt, model_id=model, **kwargs)

    raise ValueError(
        "provider must be one of: claude/anthropic, gemini/google, gpt/openai, "
        "openrouter, ollama, bedrock"
    )

# ─────────────────────────────────────────────────────────────────────────────
# AGENT FACTORY  — same signature and structure as branding_agent.py
# ─────────────────────────────────────────────────────────────────────────────

def define_strands_agent(model_id: str = "us.anthropic.claude-sonnet-5",tools_import=[],max_tokens=4096,
                        min_confidence: float = 0.70,floor_confidence: float = 0.30,sys_prompt=None,
                        persona='general',show=0) -> Agent:
    
    """
    Build and return the Dataset Discovery Agent.

    TERMINATE / FREE RESOURCES: del agent
    'anthropic.claude-fable-5'
    'anthropic.claude-opus-4-8'
    'us.anthropic.claude-opus-4-7'
    'us.anthropic.claude-sonnet-4-6'          - WORKS US & UK
    'us.anthropic.claude-opus-4-6-v1'          - WORKS US & UK
    'us.anthropic.claude-opus-4-5-20251101-v1:0'  - WORKS US & UK
    'us.anthropic.claude-sonnet-4-5-20250929-v1:0' - WORKS US & UK
    'us.anthropic.claude-haiku-4-5-20251001-v1:0'  - WORKS US & UK
    """
    region_name = "eu-central-1" if model_id[:2] == "eu" else "us-east-1"
    boto3.setup_default_session(region_name=region_name)

    model = BedrockModel(model_id=model_id,max_tokens=max_tokens)

    # --- create the agent with tools -------------------------------------------
    # Check environment type - AWS Sagemaker vs local
    aws=0
    try:
        if "HOME" in os.environ:
            if "ec2-user" in os.environ.get("HOME"):
                aws=1
                if show: print('********** Running on AWS')
    except:
        if show: print('********** Running on Locally')

    # if on AWS:
    recommended_tools= agt.agentic_toolsets(persona)
    if aws:
        from strands_tools import python_repl, file_read, file_write, current_time, http_request, diagram
        recommended_tools = recommended_tools+[python_repl, file_read, current_time, http_request, diagram]

        '''
        recommended_tools = [crawl_site, identify_datasets, extract_dataset, normalize_to_csv,
                             request_human_review, fetch_page, save_file, save_scraper_code,
                             run_python]
        '''
    ALL_TOOLS = tools_import+recommended_tools

    if sys_prompt is None: 
        sys_prompt=(
            f"You discover and download datasets from a website. Workflow:\n"
            f"(1) crawl_site on the seed URL. Never go UP in HTML page hierarchy, "
            f"only go DOWN or to linked pages;\n"
            f"(2) identify_datasets against the user's request;\n"
            f"(3) for each candidate with confidence >= {min_confidence} "
            f"call extract_dataset then normalize_to_csv;\n"
            f"(4) for confidence in [{floor_confidence}, {min_confidence}) you MUST call "
            f"request_human_review first and only extract if approved;\n"
            f"(5) ignore anything below {floor_confidence}. Never extract a "
            f"mid-confidence dataset without an approval.\n"
            f"(6) IMPORTANT: Whenever you develop and validate a FINAL, working "
            f"web-scraping Python script for the given use case, you MUST call "
            f"save_scraper_code to persist that final code to a uniquely named "
            f".py file for future reuse and reference. Always do this before "
            f"finishing the task. Pass a meaningful 'use_case' slug, the "
            f"'source_url', and a short 'description'."
        )

    agent = Agent(
        model=model,
        tools=ALL_TOOLS,
        system_prompt=sys_prompt,
    )

    return agent

def invoke_with_continue(agent, prompt, max_continues=6, show=0):
    """
    Call the agent, automatically resuming when the model stops due to
    MaxTokensReachedException. Strands appends the partial assistant message
    to history, so re-invoking (with an empty nudge) continues generation.
    """
    try:
        response = agent(prompt)
        return response
    except MaxTokensReachedException:
        if show:
            print("   ↪ max_tokens hit — auto-continuing...")

    last_response = None
    for i in range(max_continues):
        try:
            # Empty/continuation nudge; partial message is already in history.
            last_response = agent("Continue exactly where you left off. "
                                   "Be concise; do not repeat prior output.")
            return last_response
        except MaxTokensReachedException:
            if show:
                print(f"   ↪ max_tokens hit again (continue {i+1}/{max_continues})...")
            continue

    # Exhausted continuations — return whatever we have rather than crashing.
    if show:
        print("   ⚠ Reached max continuations; returning partial result.")
    return last_response
    
def itad_crawls(pgs, user_prompt, mid="us.anthropic.claude-sonnet-5",
                min_conf=0.7, floor_conf=0.3, raise_exception=0,
                max_tokens=8192, show=0, max_continues=6):
    try:
        agent
        print('Agent already exists: deleting agent...')
        del agent
    except NameError:
        pass

    print('Creating agent...')
    agent = define_strands_agent(model_id=mid, min_confidence=min_conf,
                                 floor_confidence=floor_conf, max_tokens=max_tokens)

    if not isinstance(pgs, list):
        pgs = [pgs]

    n = len(pgs)
    count = 0
    for p in pgs:
        print(f'************* Processing URL: {p}')
        ps = p.replace('\\', '-').replace('/', '-').replace(':', '').replace('.', '-')

        if isinstance(p, str):
            try:
                response = invoke_with_continue(
                    agent,
                    f"Seed URL: {p}\n\n{user_prompt}",
                    max_continues=max_continues,
                    show=show,
                )
                if show:
                    print(response)

                fn = f'{ps}.txt'
                with open(fn, "w") as file:
                    file.write(str(response))
                count += 1
            except Exception:
                print('>>>>>>>>>>>>> RUN failed ...')
                exc_type, exc_value, exc_traceback = sys.exc_info()
                traceback.print_exception(exc_type, exc_value, exc_traceback,
                                          limit=5, file=sys.stdout)
                if raise_exception:
                    raise Exception('...stopping since raise_exception is True')

    print(f'{count} of {n} runs successfully completed')
    print('Deleting agent to free resources')
    del agent

# --------------------------------------------------------------------------- #
# Agent builder (repurposed from the dataset-discovery version)
# --------------------------------------------------------------------------- #

def build_agent(model_id: str = 'eu.anthropic.claude-sonnet-5',
                temp: float = 0.4, min_confidence: float = 0.75,
                floor_confidence: float = 0.4, seek_approval: bool = False) -> Agent:
    """An agent that explores Office files and translates their VBA into one
    Python data pipeline, model-first (it reasons about ordering and retries).

    Args:
        model_id: Bedrock model id. 'eu.' prefix routes to eu-central-1.
        temp: Sampling temperature. Lower is better for faithful code translation.
        min_confidence: At/above this, translate a module directly.
        floor_confidence: Below this, skip the module (flag it in the report).
        seek_approval: If True, modules in [floor, min) require human approval.

    EXAMPLE
    agent = build_agent(seek_approval=True)
    agent(
        "Explore the Office files in ./macro_workbooks, extract their VBA and "
        "input data, and produce ./out/pipeline.py plus CSVs in ./out/data. "
        "Give me a final report of what was translated and what was skipped."
    )
    """
    aws_region = 'eu-central-1' if model_id[:2] == 'eu' else 'us-east-1'
    boto3.setup_default_session(region_name=aws_region)   # fixed: was region_name

    model = BedrockModel(model_id=model_id, temperature=temp)

    common = f'''You explore MS Office files, extract their embedded VBA macros and
input data, and translate everything into a SINGLE Python file implementing the
equivalent data pipeline, plus CSV files holding the embedded input data.

You translate VBA to Python yourself, in your reasoning. The tools only do
deterministic work (discovery, extraction, data dump, static analysis,
assembly, syntax-check). When translating: replace cell/range/sheet access with
pandas reads of the extracted CSVs; turn each VBA Sub/Function into a Python
function; preserve the logic and the order of operations faithfully; do not
invent behaviour. Keep one function per VBA procedure and reference its origin.

Workflow:
(1) discover_office_files on the input path; process only files with has_macros
    (and dump data from spreadsheet files regardless).
(2) For each file: extract_vba to get module source, and extract_embedded_data
    to write the input CSV(s).
(3) analyze_macros to get per-module translation_confidence and the suggested
    execution order (auto_exec modules first).
(4) For each module with confidence >= {{gate}}: translate its VBA to Python.
'''

    if seek_approval:
        gate = min_confidence
        tail = f'''(5) For confidence in [{floor_confidence}, {min_confidence}) you MUST call
    request_human_review first and translate only if it returns APPROVED.
(6) Ignore anything below {floor_confidence}; list it as untranslated in the report.
(7) Call assemble_python_pipeline with all translated functions (in execution
    order), the imports, a main() body that reads the CSVs and runs the pipeline,
    and the provenance list. If syntax_ok is false, fix the code and retry.
Never translate a mid-confidence module without an approval.'''
    else:
        gate = floor_confidence
        tail = f'''(5) Ignore anything below {floor_confidence}; list it as untranslated in the report.
(6) Call assemble_python_pipeline with all translated functions (in execution
    order), the imports, a main() body that reads the CSVs and runs the pipeline,
    and the provenance list. If syntax_ok is false, fix the code and retry.'''

    system_prompt = common.replace("{gate}", str(gate)) + "\n" + tail

    return Agent(
        model=model,
        tools=agt.agentic_toolsets(subset='vba'),
        #callback_handler=loop_observer,
        system_prompt=system_prompt,
    )

# ------------------------------------------------------------------------------------------------- SPECIFIC APPLICATIONS

# ************ JSON FIXER

"""
JSON sanitization utility.

Recursively walks a Python object (dict / list / scalar) and converts any
values that would normally break json.dumps() into JSON-safe equivalents.

Covers the failure modes discussed:
    1. numpy scalar/array types (np.int64, np.float32, np.bool_, np.ndarray)
    2. Enum members
    3. Pydantic models / dataclasses / arbitrary objects with __dict__
    4. NaN / Infinity floats (silently "valid" per json.dumps but invalid
       per strict JSON / most downstream parsers)
    5. datetime / date / time objects
    6. Decimal
    7. set / frozenset / tuple
    8. bytes
    9. Non-string dict keys
   10. Circular references (raises a clear error instead of RecursionError)

from __future__ import annotations

import dataclasses
import datetime as dt
import decimal
import enum
import json
import math
from typing import Any, Callable, Optional
"""

class JSONSanitizeError(ValueError):
    """Raised when a value cannot be made JSON-safe and no fallback applies."""


def make_json_safe(
    obj: Any,
    *,
    nan_policy: str = "null",       # "null" | "string" | "raise" | "keep"
    unknown_policy: str = "str",    # "str" | "raise" | "omit"
    max_depth: int = 100,
    _seen: Optional[set] = None,
    _depth: int = 0,
) -> Any:
    """Recursively convert `obj` into something json.dumps() can serialize
    with default settings, without silently producing invalid JSON.

    Parameters
    ----------
    obj : Any
        The object to sanitize (dict, list, scalar, or arbitrary object).
    nan_policy : str
        How to handle float('nan') / inf / -inf:
          - "null"   -> replace with None (default; safest for strict JSON consumers)
          - "string" -> replace with "NaN" / "Infinity" / "-Infinity"
          - "raise"  -> raise JSONSanitizeError
          - "keep"   -> leave as-is (matches json.dumps default behavior,
                        which emits non-standard tokens)
    unknown_policy : str
        How to handle objects with no recognized conversion:
          - "str"    -> fall back to str(obj) (default; lossy but never fails)
          - "raise"  -> raise JSONSanitizeError
          - "omit"   -> drop the key/item entirely
    max_depth : int
        Recursion guard for deeply nested or malformed structures.
    _seen, _depth : internal
        Used for circular-reference detection; do not set manually.

    Returns
    -------
    A structure containing only: dict, list, str, int, float, bool, None.

    Examples
    --------
    >>> import numpy as np
    >>> from enum import Enum
    >>> class Rec(Enum):
    ...     MATCH = "MATCH"
    >>> d = {
    ...     "rating": np.int64(68),
    ...     "recommendation": Rec.MATCH,
    ...     "score": np.float32(0.947),
    ...     "checked_at": dt.datetime(2026, 7, 7, 12, 0),
    ...     "tags": {"hr", "payroll"},
    ...     "raw": b"\\x00\\x01",
    ...     "bad_float": float("nan"),
    ... }
    >>> safe = make_json_safe(d)
    >>> json.dumps(safe)  # doctest: +SKIP
    """
    if _seen is None:
        _seen = set()

    if _depth > max_depth:
        raise JSONSanitizeError(
            f"Exceeded max_depth={max_depth}; object is too deeply nested "
            "or contains a structure the recursion guard couldn't unwrap."
        )

    # --- Primitives that are already JSON-safe -----------------------------
    if obj is None or isinstance(obj, (str, bool)):
        return obj

    if isinstance(obj, int):
        return obj

    if isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return _handle_non_finite_float(obj, nan_policy)
        return obj

    # --- Circular reference guard (containers only) -------------------------
    if isinstance(obj, (dict, list, tuple, set, frozenset)):
        obj_id = id(obj)
        if obj_id in _seen:
            raise JSONSanitizeError(
                f"Circular reference detected at depth {_depth} "
                f"(object type: {type(obj).__name__})."
            )
        _seen = _seen | {obj_id}

    # --- dict ----------------------------------------------------------------
    if isinstance(obj, dict):
        result = {}
        for k, v in obj.items():
            safe_key = _sanitize_key(k)
            try:
                result[safe_key] = make_json_safe(
                    v, nan_policy=nan_policy, unknown_policy=unknown_policy,
                    max_depth=max_depth, _seen=_seen, _depth=_depth + 1,
                )
            except JSONSanitizeError:
                if unknown_policy == "omit":
                    continue
                raise
        return result

    # --- list / tuple / set / frozenset --------------------------------------
    if isinstance(obj, (list, tuple, set, frozenset)):
        out = []
        for item in obj:
            try:
                out.append(make_json_safe(
                    item, nan_policy=nan_policy, unknown_policy=unknown_policy,
                    max_depth=max_depth, _seen=_seen, _depth=_depth + 1,
                ))
            except JSONSanitizeError:
                if unknown_policy == "omit":
                    continue
                raise
        return out

    # --- Enum ------------------------------------------------------------
    if isinstance(obj, enum.Enum):
        return make_json_safe(
            obj.value, nan_policy=nan_policy, unknown_policy=unknown_policy,
            max_depth=max_depth, _seen=_seen, _depth=_depth,
        )

    # --- datetime / date / time --------------------------------------------
    if isinstance(obj, (dt.datetime, dt.date, dt.time)):
        return obj.isoformat()

    # --- Decimal -------------------------------------------------------------
    if isinstance(obj, decimal.Decimal):
        # int if it's a whole number, else float; avoids importing simplejson
        return int(obj) if obj == obj.to_integral_value() else float(obj)

    # --- bytes / bytearray ---------------------------------------------------
    if isinstance(obj, (bytes, bytearray)):
        try:
            return obj.decode("utf-8")
        except UnicodeDecodeError:
            import base64
            return base64.b64encode(bytes(obj)).decode("ascii")

    # --- numpy scalars & arrays (optional dependency, imported lazily) -------
    np_result = _try_numpy(obj, nan_policy, unknown_policy, max_depth, _seen, _depth)
    if np_result is not _SENTINEL:
        return np_result

    # --- pandas Timestamp / NaT (optional dependency) -------------------------
    pd_result = _try_pandas(obj)
    if pd_result is not _SENTINEL:
        return make_json_safe(
            pd_result, nan_policy=nan_policy, unknown_policy=unknown_policy,
            max_depth=max_depth, _seen=_seen, _depth=_depth,
        )

    # --- Pydantic models (v1 or v2) -------------------------------------------
    if hasattr(obj, "model_dump"):        # Pydantic v2
        return make_json_safe(
            obj.model_dump(), nan_policy=nan_policy, unknown_policy=unknown_policy,
            max_depth=max_depth, _seen=_seen, _depth=_depth,
        )
    if hasattr(obj, "dict") and callable(getattr(obj, "dict")):  # Pydantic v1
        try:
            return make_json_safe(
                obj.dict(), nan_policy=nan_policy, unknown_policy=unknown_policy,
                max_depth=max_depth, _seen=_seen, _depth=_depth,
            )
        except TypeError:
            pass  # not actually a pydantic-style .dict(); fall through

    # --- dataclasses -----------------------------------------------------------
    if dataclasses.is_dataclass(obj) and not isinstance(obj, type):
        return make_json_safe(
            dataclasses.asdict(obj), nan_policy=nan_policy, unknown_policy=unknown_policy,
            max_depth=max_depth, _seen=_seen, _depth=_depth,
        )

    # --- generic objects with __dict__ (last resort before string fallback) ----
    if hasattr(obj, "__dict__"):
        return make_json_safe(
            vars(obj), nan_policy=nan_policy, unknown_policy=unknown_policy,
            max_depth=max_depth, _seen=_seen, _depth=_depth,
        )

    # --- Nothing matched: apply unknown_policy ----------------------------------
    if unknown_policy == "raise":
        raise JSONSanitizeError(
            f"No JSON-safe conversion available for type {type(obj).__name__!r} "
            f"(value: {obj!r})."
        )
    if unknown_policy == "omit":
        raise JSONSanitizeError("omit-signal")  # caught by caller's container loop
    return str(obj)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

_SENTINEL = object()


def _sanitize_key(key: Any) -> str:
    """JSON object keys must be strings; coerce anything else."""
    if isinstance(key, str):
        return key
    if isinstance(key, bool):
        return "true" if key else "false"
    if isinstance(key, (int, float)):
        return str(key)
    if isinstance(key, enum.Enum):
        return str(key.value)
    return str(key)


def _handle_non_finite_float(value: float, policy: str) -> Any:
    if policy == "keep":
        return value
    if policy == "null":
        return None
    if policy == "string":
        if math.isnan(value):
            return "NaN"
        return "Infinity" if value > 0 else "-Infinity"
    if policy == "raise":
        raise JSONSanitizeError(f"Non-finite float encountered: {value!r}")
    raise ValueError(f"Unknown nan_policy: {policy!r}")


def _try_numpy(obj, nan_policy, unknown_policy, max_depth, seen, depth):
    try:
        import numpy as np
    except ImportError:
        return _SENTINEL

    if isinstance(obj, np.bool_):
        return bool(obj)
    if isinstance(obj, np.integer):
        return int(obj)
    if isinstance(obj, np.floating):
        val = float(obj)
        if math.isnan(val) or math.isinf(val):
            return _handle_non_finite_float(val, nan_policy)
        return val
    if isinstance(obj, np.ndarray):
        return make_json_safe(
            obj.tolist(), nan_policy=nan_policy, unknown_policy=unknown_policy,
            max_depth=max_depth, _seen=seen, _depth=depth,
        )
    return _SENTINEL


def _try_pandas(obj):
    try:
        import pandas as pd
    except ImportError:
        return _SENTINEL

    if obj is pd.NaT:
        return None
    if isinstance(obj, pd.Timestamp):
        return obj.isoformat()
    if isinstance(obj, pd.Series):
        return obj.tolist()
    if isinstance(obj, pd.DataFrame):
        return obj.to_dict(orient="records")
    return _SENTINEL


# --------------------------------------------------------------------------- #
# Convenience wrapper matching json.dumps' call signature
# --------------------------------------------------------------------------- #

def safe_json_dumps(
    obj: Any,
    *,
    nan_policy: str = "null",
    unknown_policy: str = "str",
    max_depth: int = 100,
    **json_dumps_kwargs: Any,
) -> str:
    """Drop-in replacement for json.dumps() that never raises
    'Object of type X is not JSON serializable' and never silently
    emits non-standard NaN/Infinity tokens (unless nan_policy='keep').

    Example
    -------
    >>> safe_json_dumps({"rating": np.int64(68)})  # doctest: +SKIP
    '{"rating": 68}'
    """
    cleaned = make_json_safe(
        obj, nan_policy=nan_policy, unknown_policy=unknown_policy, max_depth=max_depth,
    )
    return json.dumps(cleaned, **json_dumps_kwargs)


def diagnose(obj: Any, _path: str = "$") -> list[str]:
    """Walk `obj` and report every field that would fail plain json.dumps(),
    without modifying anything. Useful for finding the exact culprit before
    deciding whether to sanitize or fix it upstream at the source.

    Returns a list of strings like:
        "$.rating: numpy.int64 (68)"
        "$.checked_at: datetime.datetime (2026-07-07 12:00:00)"
    """
    problems: list[str] = []
    try:
        json.dumps(obj)
        return problems  # already fine, nothing to report
    except TypeError:
        pass

    def _walk(o: Any, path: str) -> None:
        if o is None or isinstance(o, (str, int, float, bool)):
            try:
                json.dumps(o)
            except TypeError:
                problems.append(f"{path}: {type(o).__name__} ({o!r})")
            return
        if isinstance(o, dict):
            for k, v in o.items():
                if not isinstance(k, str):
                    problems.append(f"{path}.<key {k!r}>: {type(k).__name__} (non-string key)")
                _walk(v, f"{path}.{k}" if isinstance(k, str) else f"{path}[{k!r}]")
            return
        if isinstance(o, (list, tuple, set, frozenset)):
            for i, item in enumerate(o):
                _walk(item, f"{path}[{i}]")
            return
        problems.append(f"{path}: {type(o).__name__} ({o!r})")

    _walk(obj, _path)
    return problems

'''
if __name__ == "__main__":
    # Quick self-test covering every failure mode discussed.
    import numpy as np

    class Recommendation(enum.Enum):
        MATCH = "MATCH"

    @dataclasses.dataclass
    class Nested:
        note: str
        score: "np.floating"

    sample = {
        "name": "Jasmine McCoy",
        "rating": np.int64(68),                      # numpy int
        "recommendation": Recommendation.MATCH,       # Enum
        "score": np.float32(0.947),                   # numpy float
        "bad_float": float("nan"),                    # NaN
        "checked_at": dt.datetime(2026, 7, 7, 12, 0),  # datetime
        "tags": {"hr", "payroll"},                    # set
        "raw": b"\x00\x01hello",                       # bytes
        "cost": decimal.Decimal("1234.50"),           # Decimal
        "nested": Nested(note="ok", score=np.float64(1.0)),  # dataclass w/ numpy
        123: "numeric key",                            # non-string key
    }

    print("--- diagnose() on the raw object ---")
    for problem in diagnose(sample):
        print(" ", problem)

    print("\n--- sanitized output ---")
    print(safe_json_dumps(sample, indent=2))
'''

# ************* CV ASSESSER

def merge_json_files(
    input_dir: str,
    output_file: str,
    pattern: str = ''
) -> None:
    
    """
    Read all JSON files in a directory, load them as Python dictionaries,
    combine them into a list, and write them to a single JSON file.

    :param input_dir: Directory containing JSON files
    :param output_file: Path to the output JSON file
    """

    combined_data: List[Dict] = []

    for filename in os.listdir(input_dir):
        if len(pattern)<1:
            cond = filename.lower().endswith(".json")
        else:
            cond = (filename.lower().endswith(".json")) and (pattern.lower() in filename.lower())
        
        if cond:
            file_path = os.path.join(input_dir, filename)

            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                combined_data.append(data)

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(combined_data, f, indent=2)

def extract_json_span(text: str) -> str | None:
    """Return the substring from the first '{' or '[' to the last
    matching-type '}' or ']', trimming any prose/preamble/fencing
    around it. Returns None if no bracket is found at all.
    """
    match = re.search(r'[\{\[].*[\}\]]', text, flags=re.DOTALL)
    return match.group(0) if match else None

def process_cv_data(out):

    if isinstance(out,str):
        with open(out, "r", encoding="utf-8") as f:
            out = json.load(f)
    
    df=pd.DataFrame(out)

    df2=[]
    for i in range(len(out)):
        string = extract_json_span(out[i]['response'])
        try:
            df2.append(json.loads(string))
        except Exception as e:
            print(f'JSON no {i} failed to parse: \n {e} \n \n')
            exc_type, exc_value, exc_traceback = sys.exc_info()
            traceback.print_exception(exc_type, exc_value, exc_traceback,limit=5, file=sys.stdout)
            print(f'\n \n {string}')
            
            try:
                print("\n --- diagnose() on the raw object ---")
                for problem in diagnose(string):
                    print("\n >>> ", problem)

                print("\n--- sanitized output ---")

                df2.append(json.loads(safe_json_dumps(string, indent=4)))
            except Exception as e:
                print(f'JSON no {i} failed to safe-parse: \n {e}')
        
    df2=pd.DataFrame(df2)

    dff=pd.concat([df2,df[['cv']]],axis=1)

    cs=['strengths','gaps',	'red_flags']

    for i in range(len(dff)):
        for c in cs:
            d=dff.loc[i,c]
            dstr=''
            if not isinstance(d,list):d=[d]
            for dd in d:
                
                if len(dstr)==0:
                    dstr=dd
                else:
                    dstr=dstr+'\n\n'+dd
                    
            dff.loc[i,c]=dstr
    
    fn_fd = 'merged_jsons.csv'
    dff.to_csv(fn_fd,index=None)

"""
Detect and fix formatting issues in recruiter-pipeline JSON records that bundle:
  - a raw PDF-extracted CV text ("cv")
  - an LLM-generated JSON assessment ("response")
  - source PDF metadata (embedded inside "prompt" as LangChain Document reprs)

Run: python fix_cv_json_issues.py input.json output.json
"""

import json
import re, unicodedata
import sys
from collections import Counter
import pandas as pd


# ---------------------------------------------------------------------------
# 1. Detect + fix issues in the LLM "response" field (itself a JSON string)
# ---------------------------------------------------------------------------

EXPECTED_RESPONSE_KEYS = {
    "name", "email", "phone", "location", "rating", "recommendation",
    "summary", "strengths", "gaps", "red_flags", "interview_focus_areas",
    "confidence", "confidence_rationale",
}


_VALID_ESCAPE_CHARS = set('"\\/bfnrtu')
_HEX_DIGITS = set('0123456789abcdefABCDEF')


def sanitize_json_backslashes(s: str) -> str:
    """
    Make backslash usage valid JSON without corrupting already-valid escapes.

    Works on whole runs of consecutive backslashes (not char-by-char), because
    JSON parses backslashes in pairs: a run of N backslashes is N//2 literal
    backslashes, plus (if N is odd) one leftover backslash that tries to
    escape whatever character follows the run.

      - even-length run  -> already valid, no matter what follows. Leave it.
      - odd-length run   -> the trailing backslash escapes the next char.
                             If that's a valid escape (or a real \\uXXXX),
                             leave it. Otherwise, add one backslash so the
                             run becomes even and the following char is just
                             literal text again.
    """
    if not isinstance(s, str):
        return s

    def fix_run(m):
        run = m.group(0)
        n = len(run)
        pos = m.end()
        follow = s[pos] if pos < len(s) else ''

        if n % 2 == 0:
            return run  # already valid pairing — do not touch

        if follow == 'u':
            hex_part = s[pos + 1: pos + 5]
            if len(hex_part) == 4 and all(c in _HEX_DIGITS for c in hex_part):
                return run  # valid \uXXXX
            return run + '\\'  # bogus \u -> neutralize

        if follow in _VALID_ESCAPE_CHARS:
            return run  # valid \" \\ \/ \b \f \n \r \t

        return run + '\\'  # invalid target -> make the run even

    return re.sub(r'\\+', fix_run, s)

def sanitize_json_string_newlines(s: str) -> str:
    """Escape raw newlines/tabs that appear inside JSON string values
    (outside of already-escaped sequences), which otherwise break json.loads."""
    out = []
    in_string = False
    escape = False
    for ch in s:
        if in_string:
            if escape:
                out.append(ch)
                escape = False
            elif ch == '\\':
                out.append(ch)
                escape = True
            elif ch == '"':
                out.append(ch)
                in_string = False
            elif ch == '\n':
                out.append('\\n')
            elif ch == '\t':
                out.append('\\t')
            elif ch == '\r':
                out.append('\\r')
            else:
                out.append(ch)
        else:
            out.append(ch)
            if ch == '"':
                in_string = True
    return ''.join(out)

def parse_with_duplicate_detection(json_str: str):
    """Parse JSON while surfacing any duplicate keys that json.loads would
    otherwise silently overwrite (last value wins, first is lost).

    Sanitizes malformed backslashes first, and never raises: on failure it
    returns an empty object plus a '_parse_error' entry in duplicates for
    the caller to inspect instead of crashing the whole pipeline.
    """
    duplicates = {}

    def pairs_hook(pairs):
        counts = Counter(k for k, _ in pairs)
        for key, count in counts.items():
            if count > 1:
                duplicates[key] = [v for k, v in pairs if k == key]
        return dict(pairs)

    safe_str = sanitize_json_backslashes(json_str)

    try:
        obj = json.loads(safe_str, object_pairs_hook=pairs_hook)
    except json.JSONDecodeError as e:
        duplicates["_parse_error"] = str(e)
        obj = {}

    return obj, duplicates


def validate_schema(obj: dict, expected_keys: set) -> dict:
    """Report keys that shouldn't be there, and keys that are missing."""
    return {
        "extra_keys": sorted(set(obj.keys()) - expected_keys),
        "missing_keys": sorted(expected_keys - set(obj.keys())),
    }


def looks_truncated(value: str) -> bool:
    """Heuristic: a text field that doesn't end on sentence punctuation is
    a likely truncation, not just a stylistic choice."""
    value = value.strip()
    if not value:
        return False
    return value[-1] not in ".!?\"'\u201d)"


def fix_response_object(obj: dict, expected_keys: set) -> dict:
    """Drop keys outside the schema and flag (but don't silently discard)
    truncated string fields."""
    cleaned = {k: v for k, v in obj.items() if k in expected_keys}
    truncated_fields = [
        k for k, v in cleaned.items() if isinstance(v, str) and looks_truncated(v)
    ]
    if truncated_fields:
        cleaned["_truncated_fields_flagged"] = truncated_fields
    return cleaned


# ---------------------------------------------------------------------------
# 2. Clean PDF-extraction artifacts in the "cv" text
# ---------------------------------------------------------------------------

def clean_cv_text(text: str) -> str:
    original_text = text  # fallback if regex cleanup fails partway through

    try:
        # Rejoin a hyphenated word/URL split across a line break, e.g.
        # "garry-mc\ngowan-09a406b5" -> "garry-mcgowan-09a406b5"
        text = re.sub(r"-\n(?=\w)", "-", text)

        # Bullets extracted AFTER their item text ("item text\n\u2022") should
        # lead the item instead ("\u2022 item text").
        text = re.sub(r"([^\n\u2022]+?)\n?\u2022", r"\u2022 \1", text)

        # Rejoin a bullet phrase that got line-wrapped mid-sentence
        # (lower-case continuation on the next line signals a wrap, not a new item).
        text = re.sub(r"(?<=[a-zA-Z,/])\n(?=[a-z])", " ", text)

        # Rejoin an address split across a street line and a postcode line.
        text = re.sub(r"(Place\s+\w+)\n(?=[A-Z]{1,2}\d)", r"\1, ", text)

        # Normalise stray double newlines left after the joins above.
        text = re.sub(r"\n{2,}", "\n", text)

        # Strip an inline running-header repeat of the candidate's name
        # (all-caps line injected at a page break).
        text = re.sub(r"\n[A-Z]{2,}(?: [A-Z]{2,}){1,3}\n", "\n", text)

        # Normalise inconsistent trailing punctuation on employment-type tags.
        text = re.sub(r"\bFull-Time\.\b", "Full-Time", text)
    except Exception as e:
        print(f"Error cleaning CV text: {e}")
        text = original_text  # bail out to the unmodified original, not to an undefined value

    return text.strip()


# ---------------------------------------------------------------------------
# 3. Strip metadata bloat (e.g. embedded MSIP base64 sensitivity-label blobs)
# ---------------------------------------------------------------------------

def strip_metadata_bloat(metadata: dict, max_field_len: int = 200):
    """Return (cleaned_metadata, report_of_removed_fields)."""
    bloated = {k: len(str(v)) for k, v in metadata.items() if len(str(v)) > max_field_len}
    cleaned = {k: v for k, v in metadata.items() if len(str(v)) <= max_field_len}
    return cleaned, bloated

import re, unicodedata
import pandas as pd
from collections import Counter

try:
    import ftfy
except ImportError:
    ftfy = None

# ── 1. mojibake repair ────────────────────────────────────────────────────────
# markers: Ã / Â / â€ / Ð followed by a byte that only appears in mis-decoded UTF-8
MOJIBAKE_HINT = re.compile(
    r'Ã[\x80-\xbf\u0152\u0153\u0160\u0161\u017d\u017e\u0178\u0192\u2013\u2014'
    r'\u2018\u2019\u201c\u201d\u201a\u201e\u2020\u2021\u2022\u2026\u2030\u20ac\u2122]'
    r'|â€|â„¢|Â[\xa0-\xbf]|Ð[\x80-\xbf]|Ñ[\x80-\xbf]|ï»¿')

def _unmojibake(x: str, max_passes: int = 3) -> str:
    if ftfy is not None:
        return ftfy.fix_text(x)
    for _ in range(max_passes):
        if not MOJIBAKE_HINT.search(x):
            break
        for enc in ('cp1252', 'latin-1'):
            try:
                y = x.encode(enc).decode('utf-8')
            except (UnicodeEncodeError, UnicodeDecodeError):
                continue
            if y != x:
                x = y
                break
        else:
            break          # no encoding round-tripped -> stop
    return x

# ── 2. explicit character folding (meaning-preserving) ───────────────────────
CHAR_MAP = {
    # bullets / list markers -> line break (they usually *are* list items)
    **{c: '\n' for c in '•·‣⁃∙◦‧●○◘◙▪▫■□◾◽▸►▶➤➢❖✦'},
    # arrows -> '->'
    **{c: ' -> ' for c in '→⇒⟶➔➞➡'},
    # dashes / hyphens -> '-'
    **{c: '-' for c in '‐‑‒–—―−➖﹘﹣－'},
    # single quotes / primes -> "'"
    **{c: "'" for c in '‘’‚‛′‵´`'},
    # double quotes -> '"'
    **{c: '"' for c in '“”„‟″‴«»❝❞'},
    # misc
    '…': '...', '‰': '%', '⁄': '/', '∕': '/', '∼': '~', '≈': '~',
    '™': '', '®': '', '©': '(c)', '№': 'No.', '§': 'Section ',
    '†': '', '‡': '', '¶': '\n', '✓': 'Y', '✔': 'Y', '✗': 'N', '✘': 'N',
    '＂': '"', '＇': "'",
}
CHAR_MAP_RE = re.compile('[' + re.escape(''.join(CHAR_MAP)) + ']')

# ── 3. catch-all for whatever is left (emoji, dingbats, private use, …) ──────
DROP_CATS    = {'Cc', 'Cf', 'Co', 'Cs', 'Cn'}          # controls / format / unassigned
TO_SPACE_CATS = {'So', 'Sk', 'Zl', 'Zp', 'Zs'}         # symbols-other, modifiers, seps

def _fold_leftovers(x: str, drop_emoji: bool = True) -> str:
    buf = []
    for ch in x:
        if ch == '\n':
            buf.append(ch)
            continue
        cat = unicodedata.category(ch)
        if cat in DROP_CATS:
            continue
        if drop_emoji and cat in TO_SPACE_CATS:
            buf.append(' ')
            continue
        buf.append(ch)
    return ''.join(buf)

def _to_ascii(x: str) -> str:
    """Optional hard mode: é->e, ñ->n, anything else dropped."""
    return (unicodedata.normalize('NFKD', x)
            .encode('ascii', 'ignore').decode('ascii'))

# ── existing patterns (unchanged except noted) ───────────────────────────────
LITERAL_ESCAPES = re.compile(r'\\(?:u\{[0-9A-Fa-f]{1,6}\}|u[0-9A-Fa-f]{4}|U[0-9A-Fa-f]{8}'
                             r'|x[0-9A-Fa-f]{2}|[0-7]{1,3}|[abfrtv0])')
LITERAL_NL = re.compile(r'\\n')
BREAKS     = re.compile(r'\r\n|[\r\v\f\x85\u2028\u2029]')
SPACES     = re.compile(r'[\t\xa0\u1680\u2000-\u200a\u202f\u205f\u3000]+')
INVISIBLE  = re.compile(
    r'[\x00-\x08\x0e-\x1f\x7f-\x9f\xad\u034f\u061c\u200b-\u200f\u202a-\u202e'
    r'\u2060-\u2064\u2066-\u206f\ufe00-\ufe0f\ufeff\ufff9-\ufffb'
    r'\ufdd0-\ufdef\ufffe\uffff'
    r'\U0001d173-\U0001d17a\U000e0000-\U000e0fff]')
MULTISPACE      = re.compile(r' {2,}')
SPACE_AROUND_NL = re.compile(r'[ \t]*\n[ \t]*')
NL_RUNS         = re.compile(r'\n{3,}')
FORMULA_LEAD    = re.compile(r'^[=+\-@|%]')
MAX_CELL        = 32767


def sanitize_series(s: pd.Series, normalize='NFKC', fix_mojibake=True,
                    literal_newline='keep', newline='\n',
                    fold_chars=True, drop_emoji=True, ascii_only=False) -> pd.Series:
    """
    normalize      : 'NFC' | 'NFKC' | None.  NFKC also folds ﬁ->fi, ①->1, full-width forms.
    fix_mojibake   : repair â€¢ / Ã© style double-encoding before anything else
    fold_chars     : apply CHAR_MAP (bullets->\\n, curly quotes->ASCII, …)
    drop_emoji     : remaining symbol-class chars (emoji, dingbats) -> space
    ascii_only     : nuclear option, transliterate to pure ASCII
    """
    out = s.astype('string')

    if fix_mojibake:
        out = out.map(_unmojibake, na_action='ignore')

    if normalize:
        out = out.map(lambda x: unicodedata.normalize(normalize, x), na_action='ignore')

    if literal_newline == 'convert':
        out = out.str.replace(LITERAL_NL, '\n', regex=True)
    elif literal_newline == 'strip':
        out = out.str.replace(LITERAL_NL, '', regex=True)

    out = (out.str.replace(LITERAL_ESCAPES, '', regex=True)
              .str.replace(BREAKS, '\n', regex=True))

    if fold_chars:
        out = out.str.replace(CHAR_MAP_RE, lambda m: CHAR_MAP[m.group()], regex=True)

    out = (out.str.replace(SPACES, ' ', regex=True)
              .str.replace(INVISIBLE, '', regex=True))

    out = out.map(lambda x: _fold_leftovers(x, drop_emoji), na_action='ignore')
    if ascii_only:
        out = out.map(_to_ascii, na_action='ignore')

    out = (out.str.replace(MULTISPACE, ' ', regex=True)
              .str.replace(SPACE_AROUND_NL, '\n', regex=True)
              .str.replace(NL_RUNS, '\n\n', regex=True))

    if newline != '\n':
        out = (out.str.replace('\n', newline, regex=False)
                  .str.replace(MULTISPACE, ' ', regex=True))

    out = out.str.strip()

    hit = out.str.match(FORMULA_LEAD).fillna(False)
    out = out.mask(hit, "'" + out)

    return out.str.slice(0, MAX_CELL)


def sanitize_df(df: pd.DataFrame, copy=True, **kw) -> pd.DataFrame:
    df = df.copy() if copy else df
    for col in df.select_dtypes(include=['object', 'string']).columns:
        df[col] = sanitize_series(df[col], **kw)
    hdr_kw = {**kw, 'literal_newline': 'strip', 'newline': ' '}
    df.columns = pd.Index(sanitize_series(pd.Series(df.columns.astype(str)), **hdr_kw))
    return df


# ── diagnostic: decide policy from your actual data, don't guess ─────────────
def audit_chars(df: pd.DataFrame, top=60) -> pd.DataFrame:

    '''
    audit_chars(df).head(60)          # 1. see what's actually in there
    df = sanitize_df(df)              # 2. sensible defaults
    df = sanitize_df(df, ascii_only=True, drop_emoji=True)   # 3. if you need pure ASCII
    '''

    c = Counter()
    for col in df.select_dtypes(include=['object', 'string']).columns:
        for v in df[col].dropna().astype(str):
            c.update(ch for ch in v if ord(ch) > 127 or unicodedata.category(ch)[0] == 'C')
    rows = [{'char': ch, 'codepoint': f'U+{ord(ch):04X}', 'count': n,
             'category': unicodedata.category(ch),
             'name': unicodedata.name(ch, '<unnamed>')}
            for ch, n in c.most_common(top)]
    return pd.DataFrame(rows)

# ---------------------------------------------------------------------------
# 4. End-to-end pipeline
# ---------------------------------------------------------------------------

def process_record(record):

    if isinstance(record, dict):
        records = [record]
    elif isinstance(record, list):
        records = record.copy()
    else:
        raise TypeError(
            f"process_record expected a dict or list of dicts, got {type(record)}"
        )

    records_out = []
    reports_out = []

    for rec in records:

        report = {}

        # --- cv text ---
        if "cv" in rec and isinstance(rec["cv"], str):
            rec["cv"] = clean_cv_text(rec["cv"])

        # --- response JSON string ---
        if "response" in rec and isinstance(rec["response"], str):
            obj, duplicates = parse_with_duplicate_detection(rec["response"])
            schema_report = validate_schema(obj, EXPECTED_RESPONSE_KEYS)
            fixed_obj = fix_response_object(obj, EXPECTED_RESPONSE_KEYS)

            rec["response"] = json.dumps(fixed_obj, indent=2)

            report["response_duplicates_found"] = duplicates
            report["response_schema_issues"] = schema_report

        records_out.append(rec)
        reports_out.append(report)

    return records_out, reports_out


# ---------------------------------------------------------------------------
# 5. CLI entry point
# ---------------------------------------------------------------------------

def main(input_path, output_path):

    with open(input_path, "r", encoding="utf-8") as f:
        raw = f.read()

    try:
        sanitized = sanitize_json_backslashes(raw)
        sanitized = sanitize_json_string_newlines(sanitized)
        data = json.loads(sanitized)
    except json.JSONDecodeError as e:
        print(f"Failed to parse input file even after sanitizing backslashes: {e}")
        print("JSONDecodeError:", e.msg)
        print(f"  line {e.lineno}, column {e.colno} (char offset {e.pos})")
        # Show the actual problematic snippet
        start = max(0, e.pos - 100)
        end = min(len(sanitized), e.pos + 100)
        print("---- context around error ----")
        print(sanitized[start:end])
        print("-------------------------------")
        data = None
        return

    records_out, reports_out = process_record(data)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(records_out, f, indent=2)

    with open('errors_json.json', "w", encoding="utf-8") as f:
            json.dump(reports_out, f, indent=2)

    issues_found = sum(1 for r in reports_out if r)
    print(f"Processed {len(records_out)} record(s); {issues_found} had reportable issues.")
    for i, rep in enumerate(reports_out):
        if rep:
            print(f"  record[{i}]: {rep}")

    return(records_out)

def compare_docs(fns,ref_doc,pref='',creds=None,prompt=None,template_body=None,provider='bedrock',
                 model='us.anthropic.claude-sonnet-4-20250514-v1:0',max_tokens=2000, hlink = None,
                 temp=0.1,fn_out='CV-Assess',max_fails=1,async_api=1,test=0):
    '''
    EXAMPLE:
    
    # Folder with CVs
    pth_demo='C:\\Users\\josef.trchalik\\OneDrive - Ivy Technology\\Team\\Recruitment\\CVs\\HR demo\\'
    # Job spec / job description
    ref_demo='C:\\Users\\josef.trchalik\\OneDrive - Ivy Technology\\Team\\Recruitment\\Senior HR Generalist_ job description.docx'
    # Output filename
    fn_out_demo='comparison_results_DEMO_03Feb25.json'
    
    demo_json,demo_df=compare_docs(pth_demo,ref_demo,fn_out=fn_out_demo,test=0)
    
    TEMP: 0 = extremely accurate, 1=more creative & varied
    PROVIDER: ollama, openai, aws / bedrock / amazon, bedrock_api
    '''
    
    ref_doc_data,ref_doc_data_str=load_document_v2(ref_doc)
    
    # Iterate over the documents, load & process
    
    if isinstance(fns,str):
        print(f'Loading all files from folder {fns}')
        pref=fns
        fns = os.listdir(fns)
        fns = [f for f in fns if '.' in f]
    
    out=[]
    
    if test: 
        fns=fns[0:test]
        print(fns)
    
    fn_out = fn_out.split('.')[0]
    if 'CV_Assess' not in fn_out: fn_out= 'CV_Assess_'+fn_out
    fn_out_now = fn_out+'_'+now

    count=0
    fail_count=0

    for f in fns:
        
        f1=pref+f
        print(f'Processing file {f1}')
        a='ERROR'
        try:

            # Load the doc
            doc_data,doc_data_str=load_document_v2(f1)

            sys_prompt = '''You are an expert technical recruiter and talent assessment specialist with 15+ years of experience evaluating candidates across engineering, data science, and technology roles. Your task is to produce a rigorous, evidence-based assessment of how well a candidate's CV matches a given job specification.

                        ## INPUTS
                        You will receive two documents:
                        1. <job_specification> — The role requirements, responsibilities, and qualifications
                        2. <candidate_cv> — The candidate's curriculum vitae

                        ## ASSESSMENT METHODOLOGY

                        ### General approach
                        - Read the job specification carefully and extract all explicit requirements related to the role.
                        - Analyze each CV’s descriptions of skills, projects, tools, methods, results, and seniority indicators.
                        - Give especially high weight to concrete, hands-on experience with modern methods, complete projects, and measurable outcomes.
                        - Assign lower weight to list of skills or tools that are provided without matching, specific work experience
                        - Assign higher weight to work experience in industry and experience with duration longer than 12 months
                        - Assign lower weight to university projects and experience with duration shorter than 12 months
                        - Ensure scores are consistent and justified by evidence from the CV text.

                        Follow this structured reasoning process internally before producing output:

                        ### Step 1: Requirements Extraction
                        Parse the job specification into discrete, atomic requirements:
                        - **Must-have** (hard requirements): skills, qualifications, years of experience, certifications, legal/security requirements
                        - **Should-have** (strongly preferred): domain expertise, tooling, methodologies
                        - **Nice-to-have** (differentiators): bonus skills, secondary languages, adjacent experience
                        - **Soft requirements**: communication, leadership, collaboration signals

                        ### Step 2: Evidence Mapping
                        For each extracted requirement, search the CV for direct or indirect evidence:
                        - **Direct match**: Explicit mention with context (role, duration, outcome)
                        - **Indirect match**: Transferable or adjacent experience (e.g., similar and applicable industries, tools or methods)
                        - **Inferred match**: Reasonable inference from seniority, project scope, or industry (flag clearly as inferred)
                        - **Gap**: No evidence found

                        ### Step 3: Scoring
                        Apply weighted scoring per requirement category:
                        - Must-have: 50% of total weight
                        - Should-have: 30%
                        - Nice-to-have: 15%
                        - Soft requirements: 5%

                        Within each category, score requirements 0–5:
                        - 5: Strong, recent, demonstrated with outcomes
                        - 4: Clear evidence with moderate depth
                        - 3: Present but limited in depth, recency, or scope
                        - 2: Tangential or outdated (>5 years)
                        - 1: Weak inference only
                        - 0: No evidence

                        ### Step 4: Red Flag Detection
                        Identify potential concerns (do not over-weight, but surface them):
                        - Unexplained gaps >6 months
                        - Pattern of short tenures (<12 months across multiple roles)
                        - Inconsistencies between stated seniority and scope of work
                        - Missing mandatory credentials (clearance, license, degree if legally required)
                        - Overstated claims (e.g., "expert in X" with no supporting projects)

                        ## OUTPUT FORMAT

                        Respond with valid JSON matching this schema exactly:

                        {
                        "name": "Candidate's given name, all middle names and surname (family name)",
                        "email": "Email address if stated on the CV",
                        "phone": "Telephone number if stated on the CV, including international dialling code (e.g. 0044 (0) XXXX XXX XXX for UK)",
                        "location": "If a residential address is present, return 'state/county, country' (e.g. 'Texas, USA' or 'Surrey, UK'). If absent, return null."
                        "rating": "Overall fit score: <integer 0-100>",
                        "recommendation": "<STRONG_MATCH | MATCH | PARTIAL_MATCH | WEAK_MATCH | NO_MATCH>",
                        "summary": "Executive summary: <2-3 sentence assessment suitable for a hiring manager>",
                        "requirement_analysis": [
                            {
                            "requirement": "<verbatim or paraphrased requirement>",
                            "category": "<must_have | should_have | nice_to_have | soft>",
                            "evidence": "<specific CV excerpt or 'No evidence found'>",
                            "match_type": "<direct | indirect | inferred | gap>",
                            "score": <0-5>,
                            "notes": "<brief justification>"
                            }
                        ],
                        "strengths": ["<bullet>", "..."],
                        "gaps": ["<bullet>", "..."],
                        "red_flags": ["<bullet>", "..."],
                        "interview_focus_areas": [
                            {
                            "area": "<topic to probe>",
                            "rationale": "<why this matters for the role>",
                            "suggested_questions": ["<question 1>", "<question 2>"]
                            }
                        ],
                        "confidence": "<HIGH | MEDIUM | LOW>",
                        "confidence_rationale": "<why — e.g., CV detail quality, ambiguity in JD>"
                        }

                        ## CRITICAL RULES

                        1. **Evidence-based only**: Every claim must cite specific CV content. Never fabricate experience.
                        2. **Distinguish inference from fact**: If you infer a skill from context, label it explicitly.
                        3. **No protected-characteristic reasoning**: Do not consider or mention age, gender, ethnicity, nationality, religion, marital status, disability, or photos. If the CV contains such information, ignore it entirely. Base assessment solely on skills, experience, education, and demonstrated competencies.
                        4. **Recency matters**: Weight recent experience (<3 years) higher than older experience for rapidly evolving skills.
                        5. **Context over keywords**: A candidate who *delivered* tangible outcomes in certain area outranks one who *listed* something as a skill. Avoid keyword-matching bias.
                        6. **Calibrate confidence**: If the JD is vague or the CV lacks detail, set confidence to MEDIUM or LOW and explain why.
                        7. **Be specific in gaps**: "Lacks experience in tool / method XYZ" is useful; "Missing some skills in the field of ABC" is not.
                        8. **No hallucinated quotes**: Only quote text that appears verbatim in the CV.
                        9. **Neutral tone**: Professional, evidence-led. Avoid superlatives unless strongly justified.
                        10. **JSON only**: No prose before or after the JSON object. No markdown code fences unless your downstream parser expects them.

                        ## EDGE CASES

                        - **Career changers**: Weight transferable skills and recent reskilling (bootcamps, certs, projects) appropriately; do not penalize purely for non-linear paths.
                        - **Contractors/consultants**: Multiple short engagements are normal; do not flag as job-hopping without other signals.
                        - **Senior candidates with sparse CVs**: Senior professionals often list less detail. Note this and reduce confidence rather than penalizing.
                        - **Over-qualified candidates**: Flag as a retention/compensation risk in red_flags, not as a gap.
                        - **Ambiguous JD**: If the job spec is unclear on a requirement, note it in confidence_rationale and assess against a reasonable interpretation.
                        '''

            if prompt is None: 
                
                prompt="""You are an expert technical recruiter specialising in enterprise sales leadership for the data centre and hardware integration industry. You have deep knowledge of go-to-market strategy for hyperscale, OEM, and enterprise infrastructure customers, and you understand what separates a high-performing VP of Sales from a strong individual contributor.

                            Your task is to evaluate each candidate CV against the job specification below and assign a suitability rating from {bad} (no meaningful match) to {good} (outstanding match).

                            <evaluation_criteria>
                            When comparing the CV with the job specification, consider following:

                            1. Domain-specific experience (highest weight):
                            - Direct experience in the roles / industries listed in the job specification
                            - Proven track record(s) in relevant roles
                            - Understanding of methodologies and technologies underpinning the roles / industries listed in the job specification

                            2. Business impact:
                            - Quantified business impact
                            - Evidence of ability to learn quickly, drive to self-improve, including additional training and online courses
                            - Rewards and recognition
                            - Evidence of drive to go above and beyond

                            3. Seniority and general qualifications:
                            - Years at required level
                            - Cross-functional collaboration with other relevant teams
                            - Relevant education and credentials

                            Use the following weighting rules for assessment:
                            - Give higher weight to candidates with relevant hands-on experience in related fields over candidates whose CVs only list these as industries/roles touched.
                            - Give higher weight to roles longer than 12 months and to enterprise/industry experience over short-tenure, early-career roles or academia experience.
                            - Give lower weight to skill lists unsupported by specific, quantified information or evidence.
                            - Generic experience in unrelated sectors should score lower than direct match with the job specification.
                            </evaluation_criteria>

                            <output_format>
                            Respond with valid JSON matching this schema exactly:

                            {
                            "name": "Candidate's given name, all middle names and surname (family name)",
                            "email": "Email address if stated on the CV",
                            "phone": "Telephone number if stated on the CV, including international dialling code (e.g. 0044 (0) XXXX XXX XXX for UK)",
                            "location": "If a residential address is present, return 'state/county, country' (e.g. 'Texas, USA' or 'Surrey, UK'). If absent, return null."
                            "rating": "Overall fit score: <integer 0-100>",
                            "recommendation": "<STRONG_MATCH | MATCH | PARTIAL_MATCH | WEAK_MATCH | NO_MATCH>",
                            "summary": "Executive summary: <2-3 sentence assessment suitable for a hiring manager>",
                            "strengths": ["<bullet>", "..."],
                            "gaps": ["<bullet>", "..."],
                            "red_flags": ["<bullet>", "..."],
                            "confidence": "<HIGH | MEDIUM | LOW>",
                            "confidence_rationale": "<why — e.g., CV detail quality, ambiguity in JD>"
                            }
                            </output_format>

                            <instructions>
                            - Base every claim on evidence visible in the CV. Do not infer capabilities that are not stated.
                            - Be specific: prefer specific, quantified achievements over unquantified experience or awareness
                            - Be consistent across candidates: the same evidence should yield the same score regardless of ordering.
                            - Return only the JSON object. No preamble, no markdown fences, no trailing commentary.
                            </instructions>
                            """

            query=f'''{prompt} 
            <job_specification>
            {ref_doc_data}
            </job_specification>

            <candidate_cv>
            {doc_data}
            </candidate_cv>
            '''

            show=0
            if test: 
                print(query)
                show=1
                
            '''
            llm_master(query,provider='ollama',model_code=None,temp=0.95,top_p=0.9,max_tokens=1000,odb_creds=None,show=0,web_search=0,sys_prompt=None)

            TEMP For bedrock:
            Low values (near 0.0) produce very predictable, focused, and consistent results.
            High values (near 1.0) increase the chance of varied phrasing, novel ideas, or unexpected responses—at the expense of predictability
            '''

            #r,a=llm_master(query,provider=provider,model_code=model,temp=temp,max_tokens=max_tokens,odb_creds=creds,show=test,sys_prompt=sys_prompt)
            r,a=genai_master(query,provider=provider,model_code=model,temp=temp,max_tokens=max_tokens,odb_creds=None,show=show,web_search=0,
                            sys_prompt=sys_prompt,async_api=async_api)
            
            if test: 
                print('>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> RESPONSE')
                print(r)
                print('>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> BODY')
                print(a)
            
            # Save data
            location='N/A'
            a=a.replace('json','').replace('```','')
            fail_count=0
        except Exception as e:
            fail_count+=1
            print(f'!!!! ERROR PROCESSING FILE {f} - consecutive fail count = {fail_count}')
            print(f'ERROR details: {e}')
            error_details=e

        if fail_count>=max_fails: raise Exception(f'Failed to get response from LLM after {max_fails+1} attempts. \n \n ERROR: {error_details}')

        try:
            a_dict=json.loads(a)
            if 'location' in a_dict.keys(): location=a_dict['location']
        except Exception as e:
            print(f'Conversion of ANSWER to DICT failed: {e} @ {a}')

        cv_text=''
        for di in doc_data: 
            try:
                cv_text=cv_text+'\n\n'+di.page_content
            except:
                cv_text=cv_text+'\n\n'+str(di)
        
        outi={'file':f, 'cv': cv_text, 'sys_prompt':sys_prompt, 'prompt':query, 'response':a, 'location': location, 'ref_doc':ref_doc,'model':model,'temp':temp}
        out.append(outi)
        
        print('***************************')
          
        fn_out_f = f'{f.split('.')[0]}_{count}.json'

        if test: print(outi)
            
        with open(fn_out_f, 'w') as file:
            json.dump(outi, file, indent=4)
            
        with open(f'{fn_out_now}.json', 'w') as file:
            json.dump(out, file, indent=4)
        
        time.sleep(1)
        
        count+=1

    # Convert to CSV
    #out_df=get_df_from_llm([fn_out],hlink=hlink)
    path_in=f'{fn_out_now}.json'
    path_out=f'{fn_out_now}_F.json'

    out=main(path_in, path_out)

    df=pd.DataFrame(out)

    df2=[]
    for i in range(len(out)):
        try:
            df2.append(json.loads(out[i]['response']))
        except Exception as e:
            print(f'>>>>>>>>>>>>>>>>>> ERROR: {e} \n \n {out[i]['response']}')

    df2=pd.DataFrame(df2)

    dff=pd.concat([df2,df[['cv']]],axis=1)

    cs=['strengths','gaps',	'red_flags']

    for i in range(len(dff)):
        for c in cs:
            d=dff.loc[i,c]
            if not isinstance(d,list):
                d=[d]
            dstr=''
            for dd in d:
                
                if len(dstr)==0:
                    dstr=dd
                else:
                    dstr=dstr+'\n\n'+dd
                    
            dff.loc[i,c]=dstr

    #for c in dff.columns: dff[c] = dff[c].astype(str).str.replace(r"[^\w\s\n\-.,;:!?()&@#$%*/+=\'\"—–]", '', regex=True)

    audit_chars(dff).head(60)          # 1. see what's actually in there
    #dff = sanitize_df(dff)              # 2. sensible defaults
    dff = sanitize_df(dff, ascii_only=True, drop_emoji=True)   # 3. if you need pure ASCII

    fn_fd = fn_out.split('.')[0]+'.csv'
    dff.to_csv(fn_fd,index=None)

    return(out,dff)

def get_df_from_llm(d,patterns={'file':[],'location':[],'response':[r'\d{1,2}/10?',r'\d{1,2}/20?',r'\d{1,2} out of 10?',r'\d{1,2} out of 20?'],
                               'hlink':[],'cv':[]},
                    fn='get_df_from_llm.csv',hlink=None):
    
    def process_data(d,patterns,hlink=None):
        
        out={}
        
        for k in patterns.keys():
            
            if len(patterns[k])>0:
                out[k]=[]
                out[k+'_full']=[]
            else:
                out[k]=[]

        if isinstance(d,(list,tuple))==False: d=[d]
        
        for di in d:
            #print(di)
            for k in patterns.keys():
                if isinstance(di,dict):
                    if k in di.keys():
                        # RegEx Patterns available
                        if len(patterns[k])>0:
                            outi = []
                            for p in patterns[k]:
                                outii = re.findall(p, di[k])
                                if isinstance(outii,(list,tuple)): 
                                    if len(outii)>0: 
                                        outii=outii[0]
                                    else: 
                                        outii=''
                                        
                                if len(outii)>0: 
                                    outii = outii.lower().replace(' out of ',' / ').replace(' ','')
                                    outii=f'**{outii}**'
                                    a=re.findall(r'\*\*.{1,1}/',outii)
                                    if len(a)>0: outii=re.sub(r'^\*\*','**0',outii)
                                    outi.append(outii)
        
                            if len(outi)>0: 
                                outi=outi[0]
                            else:
                                outi=''
                                
                            out[k].append(outi)
                            out[k+'_full'].append(di[k])
        
                        # Hyperlink
                        elif (k=='hlink') & (hlink is not None):
                            
                            out[k].append(f'=HYPERLINK("{hlink}//{di['file']}","LINK")')
                        
                        # OTHER: No patterns - e.g. file
                        else:
                            out[k].append(di[k])
                    else:
                        out[k].append('N/A')
                else:
                    out[k].append('N/A')

        return(out)

    if isinstance(d,str):
        if np.logical_not((len(d)<200)&('.json' in d.lower())):
            fn=d
            d=[d]
        else:
            d=[[json.loads(d)]]
    elif isinstance(d,dict):
        d=[[d]]
    elif isinstance(d,list):
        if isinstance(d[0],dict):
            d=[d]
        elif isinstance(d[0],str):
            if np.logical_not((len(d[0])<200)&('.json' in d[0].lower())): d=[d]
                
    # INPUT: d = either a list of JSON filenames or a list of lists of dictionaries
    if isinstance(d,list):
        
        out=[]
        skip=0
        count=1
        
        for di in d:
            
            fni = f'{fn}_{count}.csv'
            
            if isinstance(di,str):
                if (len(di)<200)&('.json' in di.lower()):
                    fni=di
                    with open(di , 'r') as file:
                        did = json.load(file)
                else:
                    did=json.dumps(di)
            elif isinstance(di,(dict,list)):
                did=di
            else:
                skip=1

            if skip==0:
                
                outi=process_data(did,patterns,hlink=hlink)
                try:
                    outi=pd.DataFrame.from_dict(outi)
                except Exception as e:
                    print('*************************** ERROR')
                    print(e)
                    print(outi)
                fn2=fni.replace('.json','.csv')
                outi.to_csv(fn2,index=None)
                out.append(outi)
            else:
                print('!!! Skipping !!!')
                print(type(did))
                print(type(di))
                print(did)

            count+=1
            
    elif isinstance(d,pd.DataFrame):
        try:
            out=process_data(did,patterns)
            out=pd.DataFrame.from_dict(out)
            out.to_csv(fn,index=None)
        except Exception as e:
            print('*************************** ERROR')
            print(e)
            print(out)

    '''
    ADDITIONAL COLUMNS:
    ['Location ID',	'IL/DL', 'Role', 'Applied in SF', 'Phone Number', 'Personal Email',	'Entered into SF',	'Zip',	'Start Date',	
    'Welcome Email/Dez',	'Sent D/BGC Dez',	'Sent Offer Letter/ Dez',	'Banking Info/Matt',	'Onboarded/Matt',	'Comments']
    '''

    return(out)

if __name__ == "__main__":
   print("Executed when ran directly")
else:
   print("Executed when imported")