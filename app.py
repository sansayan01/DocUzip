import tkinter as tk
import traceback
from io import StringIO
import sys
import re
import zipfile
import os
import ctypes
from docx import Document  # Ensure this import is at the top

def preprocess_code(code):
    # Add the import statement at the beginning
    import_statement = "from docx import Document\n"
    
    # Remove '/mnt/data/' from file paths
    code = re.sub(r'/mnt/data/', '', code)
    
    return import_statement + code

def execute_code():
    code = code_input.get()
    
    if not code.strip():
        output_text.config(state=tk.NORMAL)
        output_text.insert(tk.END, "No code entered.\n")
        output_text.config(state=tk.DISABLED)
        return
    
    # Preprocess the code
    processed_code = preprocess_code(code)
    
    # Clear previous output
    output_text.config(state=tk.NORMAL)
    output_text.delete("1.0", tk.END)
    
    # Redirect stdout and stderr to capture print statements and errors
    old_stdout = sys.stdout
    old_stderr = sys.stderr
    new_stdout = StringIO()
    new_stderr = StringIO()
    sys.stdout = new_stdout
    sys.stderr = new_stderr

    # File tracking
    docx_files = []
    zip_files = []

    try:
        # Execute the processed code
        exec(processed_code)
        
        # Example: If code creates a DOCX file named 'example.docx'
        # Add the names of generated DOCX files to the list
        docx_files.append('questions.docx')  # Adjust based on actual output
        
        # Determine which type of file to provide
        if zip_files:
            # Provide ZIP file
            zip_file_name = zip_files[0]  # Assuming there's only one ZIP file
            output_text.insert(tk.END, f"ZIP file created: {zip_file_name}\n")
        elif docx_files:
            # Provide DOCX file
            docx_file_name = docx_files[0]  # Assuming there's only one DOCX file
            output_text.insert(tk.END, f"DOCX file created: {docx_file_name}\n")
        else:
            output_text.insert(tk.END, "No DOCX or ZIP files created.\n")

    except Exception as e:
        # Get the exception traceback
        traceback_str = traceback.format_exc()
        output_text.insert(tk.END, f"An error occurred:\n{traceback_str}")
    finally:
        # Restore original stdout and stderr
        sys.stdout = old_stdout
        sys.stderr = old_stderr
        
        # Get output from StringIO
        output = new_stdout.getvalue()
        error = new_stderr.getvalue()
        
        if output:
            output_text.insert(tk.END, output)
        if error:
            output_text.insert(tk.END, f"Error: {error}")

    output_text.config(state=tk.DISABLED)

def paste_text():
    # Paste text from clipboard into the code_input entry widget
    code_input.insert(tk.END, root.clipboard_get())

def clear_all():
    code_input.delete(0, tk.END)  # Clear the Entry widget

def create_popup():
    # Create a new top-level window
    global root
    popup = tk.Toplevel()
    popup.title("Python to DOCX")
    popup.geometry("500x300")  # Adjusted window size to be larger

    # Add an Entry widget for code input
    global code_input
    code_input = tk.Entry(popup, width=50, font=("Courier", 12))
    code_input.pack(pady=5, padx=10, fill=tk.X)

    # Add a Text widget for output
    global output_text
    output_text = tk.Text(popup, wrap=tk.WORD, font=("Courier", 12), height=5, state=tk.DISABLED)
    output_text.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)
    
    # Add highlighted text at the bottom
    footer_frame = tk.Frame(popup)
    footer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
    
    footer = tk.Label(footer_frame, text="Made by Sayan", font=("Arial", 10, "italic"), fg="blue")
    footer.pack(side=tk.TOP, pady=5)

    # Add Buttons in the center of the footer
    button_frame = tk.Frame(footer_frame)
    button_frame.pack(side=tk.TOP, pady=5)
    
    clear_button = tk.Button(button_frame, text="Clear All", command=clear_all)
    clear_button.pack(side=tk.LEFT, padx=5)
    
    paste_button = tk.Button(button_frame, text="Paste", command=paste_text)
    paste_button.pack(side=tk.LEFT, padx=5)
    
    execute_button = tk.Button(button_frame, text="Enter", command=execute_code)
    execute_button.pack(side=tk.LEFT, padx=5)

def minimize_console():
    if sys.platform == "win32":
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)  # 6 is SW_MINIMIZE

def main():
    global root
    minimize_console()  # Minimize the console window
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    create_popup()  # Show the popup
    root.mainloop()

# Run the application
if __name__ == "__main__":
    main()
