import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import threading
import os
import subprocess
import sys
import locale
import atexit

# Set UTF-8 as default encoding for subprocess
if sys.platform.startswith('win'):
    # Force UTF-8 encoding on Windows
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

global attachment_path
attachment_path = ""
body_window = None  # Keep track of body window globally

# Function to clear files on exit
def cleanup_files():
    files_to_clear = ['body.txt', 'attachment.txt', 'subjects.csv']
    for file in files_to_clear:
        if os.path.exists(file):
            try:
                if file.endswith('.csv'):
                    # Create empty CSV with header
                    pd.DataFrame({'subject': []}).to_csv(file, index=False)
                else:
                    # Clear text files
                    open(file, 'w', encoding='utf-8').close()
            except Exception as e:
                print(f"Error clearing {file}: {str(e)}")

# Register cleanup function
atexit.register(cleanup_files)

# Function to select the contacts CSV file
def select_csv():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        csv_entry.delete(0, tk.END)
        csv_entry.insert(0, file_path)

# Function to save sender email to gmail.csv
def save_email():
    # sender_email = email_entry.get().strip()
    # if sender_email:
    #     df = pd.DataFrame({'email': [sender_email]})
    #     df.to_csv("gmail.csv", index=False)
    #     messagebox.showinfo("Success", f"Email saved: {sender_email}")
    # else:
    #     messagebox.showwarning("Warning", "Enter an email address.")
    """Save email to CSV and delete token.json if save is successful."""
    sender_email = email_entry.get().strip()
    if sender_email:
        try:
            # Save email to CSV
            df = pd.DataFrame({'email': [sender_email]})
            df.to_csv("gmail.csv", index=False)
            
            # Delete token.json if it exists
            if os.path.exists("token.json"):
                try:
                    os.remove("token.json")
                    messagebox.showinfo("Success", f"Email saved: {sender_email}\ntoken.json file deleted.")
                except Exception as e:
                    messagebox.showwarning("Warning", f"Email saved but couldn't delete token.json: {str(e)}")
            else:
                messagebox.showinfo("Success", f"Email saved: {sender_email}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save email: {str(e)}")
    else:
        messagebox.showwarning("Warning", "Enter an email address.")

# Function to open a new window to edit the email body and subject
def open_body_editor():
    global attachment_path, body_window
    
    # If body window exists, just focus it and return
    if body_window is not None:
        try:
            body_window.focus_force()  # Force focus on existing window
            body_window.lift()         # Bring window to front
            return
        except tk.TclError:  # Window was closed
            body_window = None
    
    def save_body():
        body_text = body_textbox.get("1.0", tk.END).strip()
        subject_text = subject_entry.get().strip()
        
        if subject_text:
            df = pd.DataFrame({'subject': [subject_text]})
            df.to_csv("subjects.csv", index=False)
        else:
            messagebox.showwarning("Warning", "Enter an email subject.")
            return
        
        if body_text:
            with open("body.txt", "w", encoding="utf-8") as f:
                f.write(body_text)
            
            if attachment_path:
                with open("attachment.txt", "w", encoding="utf-8") as f:
                    f.write(attachment_path)
        else:
            messagebox.showwarning("Warning", "Email body cannot be empty.")
            return
        
        messagebox.showinfo("Success", "Email body and subject saved successfully!")
        body_window.destroy()
    
    def select_attachment():
        global attachment_path
        file_path = filedialog.askopenfilename(parent=body_window)  # Set parent window
        if file_path:
            attachment_path = file_path
            attachment_label.config(text=f"Attachment: {os.path.basename(file_path)}")
            body_window.focus_force()  # Return focus to body window
    
    def on_body_window_close():
        global body_window
        body_window = None
        body_window_local.destroy()
    
    body_window_local = tk.Toplevel(root)
    body_window = body_window_local  # Store reference globally
    body_window_local.title("Edit Email Body and Subject")
    body_window_local.geometry("500x500")
    body_window_local.transient(root)  # Set main window as parent
    body_window_local.grab_set()       # Make window modal
    
    # Handle window close event
    body_window_local.protocol("WM_DELETE_WINDOW", on_body_window_close)
    
    tk.Label(body_window_local, text="Enter Email Subject:").pack(pady=5)
    subject_entry = tk.Entry(body_window_local, width=50)
    subject_entry.pack(pady=5)
    
    # Handle subject loading with error checking
    try:
        if os.path.exists("subjects.csv"):
            subjects = pd.read_csv("subjects.csv")
            if not subjects.empty and 'subject' in subjects.columns:
                subject_entry.insert(0, subjects.iloc[0]['subject'])
    except Exception:
        pass
    
    tk.Label(body_window_local, text="Write your email body below:").pack(pady=5)
    body_textbox = scrolledtext.ScrolledText(body_window_local, width=60, height=15)
    body_textbox.pack(pady=5)
    
    # Enable text editing
    body_textbox.config(state='normal')
    
    # Load existing body text if it exists
    if os.path.exists("body.txt"):
        try:
            with open("body.txt", "r", encoding="utf-8") as f:
                body_textbox.insert(tk.END, f.read())
        except Exception:
            pass
    
    tk.Button(body_window_local, text="Select Attachment", command=select_attachment).pack(pady=5)
    attachment_label = tk.Label(body_window_local, text="No attachment selected")
    attachment_label.pack()
    
    # Load existing attachment if it exists
    if os.path.exists("attachment.txt"):
        try:
            with open("attachment.txt", "r", encoding="utf-8") as f:
                attachment_path = f.read().strip()
                if attachment_path and os.path.exists(attachment_path):
                    attachment_label.config(text=f"Attachment: {os.path.basename(attachment_path)}")
        except Exception:
            pass
    
    tk.Button(body_window_local, text="Save", command=save_body, fg="white", bg="blue").pack(pady=10)

def update_log(text):
    log_text.insert(tk.END, text)
    log_text.see(tk.END)
    root.update_idletasks()

# Function to run the email sending process
def start_sending():
    contact_file = csv_entry.get().strip()
    
    if not os.path.exists(contact_file):
        messagebox.showerror("Error", "Please select a valid contacts CSV file.")
        return
    
    if not os.path.exists("body.txt"):
        messagebox.showerror("Error", "Please write and save the email body before sending.")
        return

    if not os.path.exists("subjects.csv"):
        messagebox.showerror("Error", "Please enter and save the email subject before sending.")
        return
    
    if not os.path.exists("gmail.csv"):
        messagebox.showerror("Error", "Please save sender email first.")
        return

    # Copy contact file to contacts.csv
    pd.read_csv(contact_file).to_csv("contacts.csv", index=False)
    
    # Clear log text
    log_text.delete(1.0, tk.END)
    update_log(f"Starting email campaign with {len(pd.read_csv(contact_file))} emails.\n")
    
    # Disable the start button
    start_button.config(state=tk.DISABLED)
    
    def run_email_sender():
        try:
            # Set environment variable for UTF-8 encoding
            my_env = os.environ.copy()
            my_env["PYTHONIOENCODING"] = "utf-8"
        
            startupinfo = None
            if sys.platform.startswith('win'):
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
            process = subprocess.Popen(
                [sys.executable, "send_email.py"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                bufsize=1,
                universal_newlines=True,
                encoding='utf-8',
                env=my_env,
                startupinfo=startupinfo
            )

            # Create a function to handle stream reading
            def read_stream(stream, is_error=False):
                while True:
                    line = stream.readline()
                    if not line:
                        break
                    if is_error and "Error" in line:
                        root.after(0, update_log, f"Error: {line}")
                    elif not is_error:
                        root.after(0, update_log, line)

            # Create separate threads for stdout and stderr
            stdout_thread = threading.Thread(target=read_stream, args=(process.stdout, False))
            stderr_thread = threading.Thread(target=read_stream, args=(process.stderr, True))
        
            stdout_thread.start()
            stderr_thread.start()
        
            # Wait for both threads to complete
            stdout_thread.join()
            stderr_thread.join()
        
            # Wait for process to complete
            process.wait()
        
            # Clean up
            process.stdout.close()
            process.stderr.close()
        
            # Re-enable the start button
            root.after(0, lambda: start_button.config(state=tk.NORMAL))
        
        except Exception as e:
            root.after(0, update_log, f"Error during email campaign: {str(e)}\n")
            root.after(0, lambda: start_button.config(state=tk.NORMAL))

    # Start the email sending process in a separate thread
    thread = threading.Thread(target=run_email_sender, daemon=True)
    thread.start()

# GUI Setup
root = tk.Tk()
root.title("Email Sender GUI")
root.geometry("600x500")

# Select CSV File
tk.Label(root, text="Select Contacts CSV:").grid(row=0, column=0, padx=10, pady=5)
csv_entry = tk.Entry(root, width=50)
csv_entry.grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=select_csv).grid(row=0, column=2, padx=10, pady=5)

# Enter Email
tk.Label(root, text="Enter Sender Email:").grid(row=1, column=0, padx=10, pady=5)
email_entry = tk.Entry(root, width=50)
email_entry.grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="Save Email", command=save_email).grid(row=1, column=2, padx=10, pady=5)

# Open Body Editor
tk.Button(root, text="Write Email Body", command=open_body_editor, fg="white", bg="blue").grid(row=2, column=0, columnspan=3, pady=10)

# Start Sending Button
start_button = tk.Button(root, text="Start Sending Emails", command=start_sending, fg="white", bg="green")
start_button.grid(row=3, column=0, columnspan=3, pady=10)

# Log Display
tk.Label(root, text="Log Output:").grid(row=4, column=0, padx=10, pady=5)
log_text = scrolledtext.ScrolledText(root, width=70, height=15)
log_text.grid(row=5, column=0, columnspan=3, padx=10, pady=5)

# Status Bar
status_label = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_label.grid(row=6, column=0, columnspan=3, sticky=tk.W+tk.E, padx=10, pady=5)

# Configure grid weights
root.grid_columnconfigure(1, weight=1)
for i in range(7):
    root.grid_rowconfigure(i, weight=0)
root.grid_rowconfigure(5, weight=1)  # Make the log text area expandable

# Run the GUI
root.mainloop()