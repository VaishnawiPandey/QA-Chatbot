#!/usr/bin/env python3


# Import required libraries
import pandas as pd
import os
import argparse
from google import genai
import matplotlib.pyplot as plt

def setup_gemini_api():
    api_key_env = os.environ.get("GEMINI_API_KEY")    
    if api_key_env:
        api_key = api_key_env
    else:
        api_key = "Paste your API"
        print("Warning: Using hardcoded API key. For production, use environment variables.")
    
    client = genai.Client(api_key=api_key)
    print("Gemini API client initialized successfully")
    return client


def test_api_connection(client):
    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash", 
            contents="Explain how AI works in a few words"
        )
        print("API Test Response:", response.text)
        print("API connection successful!")
        return True
    except Exception as e:
        print(f"API connection failed: {str(e)}")
        return False
    

def load_excel_file(file_path):
    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        
        dataframes = {}
        for sheet in sheet_names:
            # Check if sheet has at least 5 rows before reading with header=4
            temp_df = pd.read_excel(excel_file, sheet_name=sheet, header=None)
            if temp_df.shape[0] < 5:
                continue 
            # Now read with header=4
            df = pd.read_excel(excel_file, sheet_name=sheet, header=4)
            if df.shape[0] == 0:
                continue  # Ignore sheets with no records
            dataframes[sheet] = df
            
        print(f"Successfully loaded Excel file with sheets: {', '.join(dataframes.keys())}")
        return dataframes
    except Exception as e:
        print(f"Error loading Excel file: {str(e)}")
        return None

def get_excel_info(dataframes):
    info = "Excel File Information:\n\n"
    for sheet_name, df in dataframes.items():
        # Skip sheets with no columns or no data
        if df.shape[1] == 0 or df.shape[0] == 0:
            continue
        info += f"=== Sheet: {sheet_name} ===\n"
        info += f"Summary: This sheet contains {df.shape[0]} rows and {df.shape[1]} columns.\n"
        info += f"Columns: {', '.join(df.columns.astype(str).tolist())}\n"
        info += "All data:\n"
        info += df.to_string() + "\n"
        info += "Column data types:\n"
        for col, dtype in df.dtypes.items():
            info += f"  - {col}: {dtype}\n"
        info += "\n" + "-"*50 + "\n\n"
    return info

def count_tokens(text):
    return len(text) // 4

def ask_question_about_excel(client, dataframes, question, chat_history=None, model="gemini-2.0-flash"):
    excel_info = get_excel_info(dataframes)
    
    # Prepare chat history context
    history_context = ""
    if chat_history:
        history_context += "Previous chat history (last 5 turns):\n"
        for idx, (q, a) in enumerate(chat_history[-5:], 1):
            history_context += f"{idx}. Q: {q}\n   A: {a}\n"
        history_context += "\n"
    
    # Construct the prompt
    prompt = f"""
    You are an AI assistant specialized in analyzing Excel data. Below is information about an Excel file with multiple sheets.

    {excel_info}

    {history_context}
    User Question: {question}

    Based on the Excel data and chat history provided, please answer the question comprehensively and accurately.
    If the answer requires calculations or data analysis, perform them and explain your process.
    If you cannot answer based on the data provided, please state so clearly.
    """
    # Print estimated input token count
    token_count = count_tokens(prompt)
    print(f"Estimated input token count: {token_count}")
    try:
        response = client.models.generate_content(
            model=model,
            contents=prompt
        )
        return response.text
    except Exception as e:
        return f"Error generating response: {str(e)}"

def excel_qa_interface(client):
    print("\n" + "="*50)
    print("Excel QA Chatbot")
    print("="*50)
    
    # Get file path from user
    file_path = input("Enter the path to your Excel file: ")
    
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
    
    # Load the Excel file
    dataframes = load_excel_file(file_path)
    
    if dataframes is None:
        return
    
    print("\nYou can now ask questions about your Excel data. Type 'exit' to quit.")
    
    chat_history = []
    while True:
        question = input("\nYour question: ")
        if question.lower() in ['exit', 'quit']:
            print("Exiting the Excel QA chatbot.")
            break
        # Pass chat_history to Gemini
        answer = ask_question_about_excel(client, dataframes, question, chat_history)
        print("\nAnswer:")
        print(answer)
        print("\n" + "-" * 50)
        # Save to chat history
        chat_history.append((question, answer))

def demonstrate_multi_sheet_qa(client, file_path):
    print("\n" + "="*50)
    print("Multi-Sheet Excel QA Demonstration")
    print("="*50)
    
    # Load all sheets from the Excel file
    dataframes = load_excel_file(file_path)
    
    if dataframes is None:
        return
    
    # Display information about the sheets
    print(f"Found {len(dataframes)} sheets in the Excel file:")
    for sheet_name, df in dataframes.items():
        print(f"- {sheet_name}: {df.shape[0]} rows × {df.shape[1]} columns")
    
    # Example questions that work across multiple sheets
    example_questions = [
        "What are the names of all sheets in this Excel file?",
        "Which sheet has the most data rows?",
        "Summarize the key information from each sheet",
    ]
    
    print("\nExample questions you could ask about multiple sheets:")
    for q in example_questions:
        print(f"- {q}")
    
    # Fixed input handling to ensure question input works properly
    try_question = input("\nWould you like to try a question about the multi-sheet data? (y/n): ")
    if try_question.lower() == 'y':
        # Use a direct input with clear prompt
        print("\nEnter your question below and press Enter:")
        question = input("> ")
        print(f"Processing question: '{question}'")
        
        # Get and display answer
        answer = ask_question_about_excel(client, dataframes, question)
        print("\nAnswer:")
        print(answer)
    else:
        print("Demo completed without asking a question.")

def main():
    """Main entry point for the script"""
    parser = argparse.ArgumentParser(description='Excel QA Chatbot with multi-sheet support')
    parser.add_argument('--file', '-f', type=str, help='Path to Excel file (optional)')
    parser.add_argument('--demo', '-d', action='store_true', help='Run in demonstration mode')
    parser.add_argument('--test', '-t', action='store_true', help='Test API connection only')
    args = parser.parse_args()
    
    # Setup the Gemini API
    client = setup_gemini_api()
    
    # Test API connection if requested
    if args.test:
        test_api_connection(client)
        return
    
    # If file path is provided and demo flag is set, run demo mode
    if args.file and args.demo:
        if os.path.exists(args.file):
            demonstrate_multi_sheet_qa(client, args.file)
        else:
            print(f"Error: File not found at {args.file}")
    elif args.file:
        if os.path.exists(args.file):
            dataframes = load_excel_file(args.file)
            if dataframes:
                print(f"Excel file loaded successfully with {len(dataframes)} sheets.")
                print("Run with --demo flag to see more functionality.")
        else:
            print(f"Error: File not found at {args.file}")
    else:
        excel_qa_interface(client)

if __name__ == "__main__":
    main()
