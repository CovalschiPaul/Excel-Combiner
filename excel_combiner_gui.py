import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import logging
import concurrent.futures
import time

# Configure logging
logging.basicConfig(filename='excel_combiner.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ExcelCombinerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel Combiner")
        self.root.geometry("700x400")
        self.root.configure(bg='#ececec')

        # Initialize variables
        self.input_dir_var = tk.StringVar()
        self.template_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()

        # Start main application window
        self.start_main_app()

    def start_main_app(self):
        """Initialize and display the main application window."""
        self.create_widgets()
        self.root.deiconify()

    def create_widgets(self):
        """Create and place all widgets in the window."""
        tk.Label(self.root, text="Excel Combiner", font=('Arial', 24, 'bold'), bg='#ececec', fg='#333333').pack(pady=(20, 10))

        main_frame = tk.Frame(self.root, bg='#ececec')
        main_frame.pack(padx=30, pady=10, fill='both', expand=True)

        # Create entry widgets
        self.create_entry_with_button(main_frame, "Input Directory:", self.input_dir_var, self.browse_directory)
        self.create_entry_with_button(main_frame, "Template File:", self.template_file_var, self.browse_template_file)
        self.create_entry_with_button(main_frame, "Output File:", self.output_file_var, self.browse_output_file)

        # Combine Files Button
        tk.Button(main_frame, text="Combine Files", command=self.run_combiner, bg='#4CAF50', fg='white', font=('Arial', 14, 'bold'), height=2, padx=20).pack(pady=(20, 10))

        # Status and Progress
        self.status_bar = tk.Label(self.root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W, bg='#ececec', fg='#333333', padx=10, font=('Arial', 10))
        self.status_bar.pack(side='bottom', fill='x', pady=(10, 10))

        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=660, mode='determinate')
        self.progress.pack(pady=(10, 20), fill='x', padx=20)

    def create_entry_with_button(self, parent_frame: tk.Frame, label_text: str, text_var: tk.StringVar, button_command):
        """Create a label, entry, and browse button."""
        row_frame = tk.Frame(parent_frame, bg='#ececec')
        row_frame.pack(pady=5, fill='x')

        tk.Label(row_frame, text=label_text, bg='#ececec', fg='#333333', font=('Arial', 12), width=15, anchor='w').pack(side='left', padx=10, pady=5)
        tk.Entry(row_frame, textvariable=text_var, width=50, font=('Arial', 12)).pack(side='left', padx=10, pady=5, fill='x', expand=True)
        tk.Button(row_frame, text="Browse...", command=button_command, bg='#4CAF50', fg='white', font=('Arial', 12), width=12).pack(side='right', padx=10, pady=5)

    def browse_directory(self):
        """Open a dialog to select the input directory."""
        directory = filedialog.askdirectory(title="Select Input Directory")
        if directory:
            self.input_dir_var.set(directory)

    def browse_template_file(self):
        """Open a dialog to select the template file."""
        file_path = filedialog.askopenfilename(title="Select Template File", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.template_file_var.set(file_path)

    def browse_output_file(self):
        """Open a dialog to select the output file."""
        file_path = filedialog.asksaveasfilename(title="Select Output File", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.output_file_var.set(file_path)

    def read_template(self):
        """Read the template Excel file and return its DataFrame and column list."""
        template_file = self.template_file_var.get()
        if not os.path.isfile(template_file):
            raise FileNotFoundError(f"Template file not found: {template_file}")
        try:
            template_df = pd.read_excel(template_file)
            return template_df, template_df.columns.tolist()
        except Exception as e:
            logging.error(f"Failed to read template file {template_file}: {e}")
            raise RuntimeError(f"Failed to read template file: {e}")

    def process_file(self, file_path: str, template_columns):
        """Read an Excel file and return a DataFrame with columns aligned to the template."""
        try:
            df = pd.read_excel(file_path)
            if not df.empty:
                df = df.reindex(columns=template_columns)
                df = df.dropna(how='all')
                return df
            return pd.DataFrame(columns=template_columns)
        except Exception as e:
            logging.error(f"Failed to process file {file_path}: {e}")
            raise RuntimeError(f"Failed to process file {file_path}: {e}")

    def combine_excel_files(self, input_directory: str, template_columns) -> pd.DataFrame:
        """Combine Excel files from the input directory into a single DataFrame."""
        df_list = []
        files = [f for f in os.listdir(input_directory) if f.endswith('.xlsx')]
        self.progress['maximum'] = len(files)
        start_time = time.time()

        for i, filename in enumerate(files):
            file_path = os.path.join(input_directory, filename)
            try:
                df = self.process_file(file_path, template_columns)
                if not df.empty:
                    df_list.append(df)
            except RuntimeError as e:
                self.update_status(f"Error processing {filename}: {e}")
            self.update_progress(i + 1, len(files), start_time)

        if df_list:
            return pd.concat(df_list, ignore_index=True)
        raise RuntimeError("No valid dataframes were found.")

    def update_progress(self, current: int, total: int, start_time: float):
        """Update progress bar and text."""
        self.progress['value'] = current
        elapsed_time = time.time() - start_time
        estimated_time = (elapsed_time / current) * total - elapsed_time if current > 0 else 0
        self.status_bar.config(text=f"Progress: {current}/{total} - Estimated time left: {int(estimated_time)}s")
        self.root.update_idletasks()

    def save_combined_data(self, combined_df: pd.DataFrame):
        """Save the combined DataFrame to an Excel file."""
        output_file = self.output_file_var.get()
        if os.path.exists(output_file):
            if not messagebox.askyesno("File Exists", "The output file already exists. Do you want to overwrite it?"):
                return
        
        try:
            combined_df.to_excel(output_file, sheet_name='Sheet1', index=False, header=True)
            self.update_status(f"Data successfully combined into {output_file}")
            messagebox.showinfo("Success", f"Data successfully combined into {output_file}")
        except Exception as e:
            logging.error(f"Failed to write combined data to {output_file}: {e}")
            raise RuntimeError(f"Failed to write combined data to {output_file}: {e}")

    def update_status(self, message: str):
        """Update the status bar with a message."""
        self.status_bar.config(text=message)
        self.root.update_idletasks()

    def run_combiner(self):
        """Run the Excel combiner process."""
        input_directory = self.input_dir_var.get()
        template_file = self.template_file_var.get()
        output_file = self.output_file_var.get()

        if not input_directory or not template_file or not output_file:
            messagebox.showwarning("Input Error", "Please select all required files and directories.")
            return

        self.update_status("Processing...")

        with concurrent.futures.ThreadPoolExecutor() as executor:
            executor.submit(self.process_combination, input_directory, template_file, output_file)

    def process_combination(self, input_directory: str, template_file: str, output_file: str):
        """Process the combination in a separate thread."""
        try:
            template_df, template_columns = self.read_template()
            combined_df = self.combine_excel_files(input_directory, template_columns)
            if not combined_df.empty:
                self.save_combined_data(combined_df)
            else:
                self.update_status("No data to save. Please check the input files.")
        except Exception as e:
            self.update_status(f"Error: {e}")
        finally:
            self.update_status("Completed")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCombinerApp(root)
    root.mainloop()
