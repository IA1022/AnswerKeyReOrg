import re
from docx import Document

# --- Core Logic Functions (File Reading) ---

def read_docx_to_string(file_path):
    """Reads all paragraphs from a DOCX file and returns them as a single string."""
    try:
        document = Document(file_path)
        full_text = []
        for para in document.paragraphs:
            # Append each paragraph followed by a reliable newline
            full_text.append(para.text.strip())
        # Join with multiple newlines to clearly separate paragraphs
        return '\n\n'.join(full_text)
    except Exception as e:
        print(f"Error reading DOCX file: {e}")
        return None

# --- Core Logic Functions (Parsing) ---

def parse_quiz_content(doc_content):
    """
    Parses the full text content using a highly tolerant regex.
    """
    if not doc_content:
        return []

    # 1. Mappings and Regex Patterns
    
    # NEW, MORE TOLERANT QUESTION REGEX:
    # It looks for "Type: Multiple choice question" followed by "Question X" 
    # and captures everything in between and after until the next question or Answer key.
    QUESTION_REGEX = r'Type: Multiple choice question[\s\S]*?Question (\d+) ([\s\S]*?)(?=Type: Multiple choice question|Answer key|$)'

    # Regex to capture question number and answer letter (e.g., 1.a)
    ANSWER_REGEX = r'(\d+)\.([a-d])'

    # Pattern to detect the start of an option (e.g., "A)", "a.", "1. ", "Ghk")
    # This will be used to split the captured body into Question Text and Options.
    # We will stick to the assumption that options start on new lines.
    OPTION_PREFIX_PATTERN = r'^\s*([a-dA-D]{1}[\.\)]|\d{1}[\.\)]|\w{3,4})$' 
    # The last part (|w{3,4}) is to catch the arbitrary options like Ghk, Fcbk, etc. from your example.

    # --- 2. Extract Answers ---

    try:
        _, answer_key_block = doc_content.split("Answer key")
    except ValueError:
        print("Warning: 'Answer key' header not found. No answers will be mapped.")
        answer_key_block = ""
    
    answer_map = {}
    for match in re.finditer(ANSWER_REGEX, answer_key_block, re.DOTALL):
        q_num = int(match.group(1))
        answer_letter = match.group(2)
        answer_map[q_num] = answer_letter 

    # --- 3. Extract and Process Questions ---

    structured_data = []

    # Use re.DOTALL (re.S) to ensure '[\s\S]' works as expected across newlines
    for match in re.finditer(QUESTION_REGEX, doc_content, re.DOTALL):
        q_num = int(match.group(1))
        body = match.group(2).strip()
        
        # 3.1. Split the body into lines
        lines = [line.strip() for line in body.splitlines() if line.strip()]

        # 3.2. Identify the separation point between question text and options
        option_start_index = -1
        # Loop backwards to find the block of 4 lines that look like options
        if len(lines) >= 4:
            for i in range(len(lines) - 4, len(lines)):
                 # Check if the line matches the generic format of a short option
                if re.match(OPTION_PREFIX_PATTERN, lines[i]):
                    option_start_index = i
                    break # Stop at the first line of the option block

        # Fallback: If no pattern is found, assume the last 4 lines are the options (as in your original example)
        if option_start_index == -1 and len(lines) >= 4:
             option_start_index = len(lines) - 4

        # 3.3. Separate question text and options
        if option_start_index != -1:
            question_text = '\n'.join(lines[:option_start_index]).strip()
            # Take the 4 lines starting from the detected index
            options = lines[option_start_index:option_start_index + 4] 
        else:
            # Everything is the question text
            question_text = '\n'.join(lines).strip()
            options = []
        
        # 3.4. Build the structured dictionary
        structured_data.append({
            "q_no": q_num,
            "question": question_text,
            "options": options,
            "answer_letter": answer_map.get(q_num, 'N/A')
        })

    return structured_data

# --- Main Execution and Formatting (Same as before) ---

def format_and_print_output(structured_data):
    """Formats and prints the data in the requested custom format."""
    print("\n" + "="*50)
    print("           ✅ Extracted Quiz Data")
    print("="*50 + "\n")
    
    for item in structured_data:
        print(f"Q{item['q_no']}. {item['question']}")
        
        option_letters = ['A', 'B', 'C', 'D']
        print("\nOptions")
        
        # Print options, stripping prefixes like "A. " or "1) "
        for i, option in enumerate(item['options']):
            if i < len(option_letters):
                 # Use a simple prefix removal regex
                 clean_option = re.sub(r'^\s*([a-dA-D]{1}[\.\)]|\d{1}[\.\)])\s*', '', option).strip()
                 print(f"  {option_letters[i]}. {clean_option}")
        
        answer_display = item['answer_letter'].lower()
        if answer_display != 'n/a':
             print(f"\nAnswer - {answer_display} (Option {answer_display.upper()})")
        else:
             print(f"\nAnswer - N/A (Missing in key)")
        print("-" * 40 + "\n")

# 🛑 CHANGE THIS TO YOUR DOCUMENT PATH
FILE_PATH = "test.docx" 

print(f"Reading content from: {FILE_PATH}...")
document_text = read_docx_to_string(FILE_PATH)

if document_text:
    structured_data = parse_quiz_content(document_text)
    format_and_print_output(structured_data)
else:
    print("Script terminated due to file reading error.")