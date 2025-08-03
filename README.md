
# AI-Powered Multi-Sheet Excel Question Answering Chatbot using Google Gemini API

> A Python-based interactive console chatbot that ingests multi-sheet Excel files and provides intelligent answers via Google's Gemini generative AI.

![Architecture](docs/architecture.png)  
*System Architecture Diagram*

---

## 🧾 Table of Contents

1. Abstract
2. Introduction
3. Project Objectives
4. System Architecture
5. Methodology
6. Implementation Details 
7. Experimentation & Evaluation
8. Security & Ethical Considerations
9. Limitations  
10. Future Work  
11. Conclusion
12. References 

---

## 📌 Abstract

This project demonstrates the design and implementation of a natural language chatbot that answers questions about the contents of Excel files with multiple sheets. Using Google Gemini's generative capabilities, the chatbot extracts structural information, builds prompts, and responds conversationally. Built in Python, it showcases how LLMs can make spreadsheet analytics more intuitive.

---

## 📘 Introduction

Multi-sheet Excel files are widely used but often hard to explore without advanced tools. This console-based chatbot lets users ask natural questions like "Which sheet has the most rows?" or "What was the highest quantity sold in May?"—and gets accurate answers using Google Gemini.

---

## 🎯 Project Objectives

- ✅ Ingest Excel files with multiple heterogeneous sheets
- ✅ Summarize sheet-level metadata (rows, columns, dtypes)
- ✅ Use Gemini LLM to generate accurate responses
- ✅ Build token-aware, chat-history–driven prompts
- ✅ Provide a simple REPL interface for Q&A

---

## 🧩 System Architecture

Excel File → Pandas → Extracted Info + Chat History → Prompt → Gemini → Answer

The chatbot runs from the command line. It loads Excel data using `pandas`, crafts a well-structured prompt using extracted metadata, and interacts with Google Gemini (gemini-2.0-flash) to answer queries.

---

## 🔧 Methodology

### 1. Data Ingestion
- Uses `pandas` to load all sheets
- Skips sheets with <5 rows
- Reads headers from row 4 (custom client format)

### 2. Token Management
- Estimates token count using simple `len(text)//4` heuristic

### 3. Prompt Engineering
Includes:
- Static role instructions
- Metadata summary (from `get_excel_info()`)
- Up to 5-turn chat history
- Current user query

### 4. Gemini API Integration
- API key via env var `GEMINI_API_KEY` or fallback
- Calls `gemini-2.0-flash` for low-latency inference

### 5. Chat Workflow
```bash
User input → load_excel_file() → get_excel_info() → ask_question_about_excel() → Gemini → Response

client = setup_gemini_api()
if test_api_connection(client):
    dfs = load_excel_file("sales_fy23.xlsx")
    answer = ask_question_about_excel(client, dfs, "Which month recorded the highest total sales?")
    print(answer)

Dependencies:
Python 3.11

pandas 2.2.2

google-generativeai 0.2.1

matplotlib 3.9.0 (for future visualization)

Security & Ethical Considerations
⚠ Avoid hardcoding API keys; use env vars or secret managers

Data privacy: anonymize sensitive Excel data

LLM hallucinations: verify results, especially on critical data

Limitations
High token usage with large workbooks

No charts or visual analytics (planned)

Repeated prompts re-send full Excel summary (no caching)

Future Work
Implement RAG using vector stores (e.g., FAISS)

Add GUI or web-based version (Flask/FastAPI)

Connect to WhatsApp bot (via Ollama or Mistral)

Generate visual summaries using matplotlib

Add unit tests and GitHub Actions for CI

Conclusion
This chatbot proves that LLMs like Gemini can transform the way we interact with structured data. By simplifying spreadsheet exploration through natural language, it opens the door to more accessible, scalable, and conversational analytics.
