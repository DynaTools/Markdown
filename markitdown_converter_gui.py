import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
from typing import List, Dict, Optional, Any

# Try importing markitdown
try:
    from markitdown import MarkItDown
except ImportError:
    # If the import fails, we'll handle it in the GUI
    pass


class MarkdownConverterApp(tk.Tk):
    """Main application for markdown conversion"""

    def __init__(self):
        super().__init__()
        
        # Configure the main window
        self.title("Markdown Converter")
        self.geometry("700x500")
        self.minsize(500, 400)
        
        # Dictionary to store paths, conversion status, and output paths
        self.files: Dict[str, Dict[str, Any]] = {}
        
        # Keep track of MarkItDown object
        self.converter: Optional[Any] = None
        
        # Output directory for batch saving
        self.output_directory: Optional[str] = None
        
        # Check if markitdown is installed
        self.check_dependencies()
        
        # Create the main UI
        self.create_widgets()
    
    def check_dependencies(self):
        """Check if required dependencies are installed"""
        try:
            # Try to initialize converter with enable_plugins, if not available, try without it
            try:
                self.converter = MarkItDown(enable_plugins=True)
            except TypeError:
                # If enable_plugins is not a valid parameter, try without it
                self.converter = MarkItDown()
            return True
        except (ImportError, NameError):
            messagebox.showerror(
                "Missing Dependencies",
                "The markitdown library is not installed. Please install it using pip:\n\n"
                "pip install 'markitdown[all]'\n\n"
                "or\n\n"
                "pip install 'markitdown[pdf,docx,xlsx]'"
            )
            return False
    
    def create_widgets(self):
        """Create all the UI widgets"""
        # Configure grid layout
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0)  # Header
        self.rowconfigure(1, weight=0)  # Instructions
        self.rowconfigure(2, weight=3)  # File list
        self.rowconfigure(3, weight=0)  # Status bar
        
        # Header frame
        header_frame = ttk.Frame(self, padding="10")
        header_frame.grid(row=0, column=0, sticky="ew")
        
        ttk.Label(
            header_frame,
            text="Markdown Converter",
            font=("Arial", 16)
        ).pack(side="left", padx=5)
        
        # Add a button to manually select files
        select_btn = ttk.Button(
            header_frame,
            text="Select Files",
            command=self.select_files
        )
        select_btn.pack(side="right", padx=5)
        
        # Add button to set output directory
        output_dir_btn = ttk.Button(
            header_frame,
            text="Set Output Directory",
            command=self.select_output_directory
        )
        output_dir_btn.pack(side="right", padx=5)
        
        # Instructions frame
        instructions_frame = ttk.Frame(self, padding="10")
        instructions_frame.grid(row=1, column=0, sticky="ew", padx=10)
        
        # Output directory label
        self.output_dir_var = tk.StringVar(value="Output Directory: Default (same as input files)")
        ttk.Label(
            instructions_frame,
            textvariable=self.output_dir_var,
            font=("Arial", 9)
        ).pack(side="bottom", anchor="w", padx=5, pady=2)
        
        ttk.Label(
            instructions_frame,
            text="Select Word, Excel, or PDF files to convert to Markdown.",
            font=("Arial", 10)
        ).pack(side="top", anchor="w", padx=5, pady=5)
        
        # Create treeview for file list
        file_frame = ttk.LabelFrame(self, text="Files", padding="10")
        file_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(0, weight=1)
        
        # Scrollbar for the treeview
        scrollbar = ttk.Scrollbar(file_frame)
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Create the treeview
        columns = ("File", "Type", "Status", "Output")
        self.file_list = ttk.Treeview(
            file_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
            yscrollcommand=scrollbar.set
        )
        scrollbar.config(command=self.file_list.yview)
        
        # Configure columns
        self.file_list.heading("File", text="File")
        self.file_list.heading("Type", text="Type")
        self.file_list.heading("Status", text="Status")
        self.file_list.heading("Output", text="Output")
        
        self.file_list.column("File", width=200, minwidth=100)
        self.file_list.column("Type", width=80, minwidth=60)
        self.file_list.column("Status", width=100, minwidth=80)
        self.file_list.column("Output", width=200, minwidth=100)
        
        self.file_list.grid(row=0, column=0, sticky="nsew")
        
        # Right-click menu for file list
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Open Output", command=self.open_output)
        self.context_menu.add_command(label="Open Output Folder", command=self.open_output_folder)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Remove", command=self.remove_file)
        
        # Bind right-click to show context menu
        self.file_list.bind("<Button-3>", self.show_context_menu)
        
        # Button frame for actions
        button_frame = ttk.Frame(self, padding="10")
        button_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=10)
        
        ttk.Button(
            button_frame,
            text="Convert All",
            command=self.convert_all_files
        ).pack(side="left", padx=5)
        
        ttk.Button(
            button_frame,
            text="Clear All",
            command=self.clear_all_files
        ).pack(side="left", padx=5)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(
            button_frame,
            textvariable=self.status_var,
            relief="sunken",
            anchor="w",
            padding=(5, 0)
        ).pack(side="right", fill="x", expand=True, padx=5)
    
    def select_files(self):
        """Open file dialog to select files"""
        filetypes = (
            ("Documents", "*.docx *.xlsx *.xls *.pdf"),
            ("Word Documents", "*.docx"),
            ("Excel Spreadsheets", "*.xlsx *.xls"),
            ("PDF Documents", "*.pdf"),
            ("All Files", "*.*")
        )
        
        files = filedialog.askopenfilenames(
            title="Select files to convert",
            filetypes=filetypes
        )
        
        if files:
            self.process_files(files)
    
    def select_output_directory(self):
        """Select a directory where all converted files will be saved"""
        directory = filedialog.askdirectory(
            title="Select Output Directory for Converted Files"
        )
        
        if directory:
            self.output_directory = directory
            self.output_dir_var.set(f"Output Directory: {directory}")
            
            # Update output paths for existing files
            for file_path, info in self.files.items():
                filename = os.path.basename(file_path)
                new_output_path = os.path.join(directory, os.path.splitext(filename)[0] + ".md")
                info["output"] = new_output_path
                
                # Update the treeview
                for item in self.file_list.get_children():
                    if file_path in self.file_list.item(item, "tags"):
                        values = self.file_list.item(item, "values")
                        new_values = (values[0], values[1], values[2], os.path.basename(new_output_path))
                        self.file_list.item(item, values=new_values)
    
    def process_files(self, file_paths: List[str]):
        """Process files that were selected"""
        # Check if the converter is available
        if not self.converter and not self.check_dependencies():
            return
        
        for path in file_paths:
            # Skip files already in the list
            if path in self.files:
                continue
            
            # Get file extension and validate
            file_ext = os.path.splitext(path)[1].lower()
            if file_ext not in (".docx", ".xlsx", ".xls", ".pdf"):
                messagebox.showwarning(
                    "Unsupported File",
                    f"The file '{os.path.basename(path)}' is not supported.\n"
                    "Only Word (.docx), Excel (.xlsx/.xls), and PDF (.pdf) files are supported."
                )
                continue
            
            # Generate output path based on whether an output directory is set
            if self.output_directory:
                output_path = os.path.join(
                    self.output_directory, 
                    os.path.splitext(os.path.basename(path))[0] + ".md"
                )
            else:
                output_path = os.path.splitext(path)[0] + ".md"
            
            # Store file info
            self.files[path] = {
                "path": path,
                "name": os.path.basename(path),
                "type": file_ext[1:].upper(),
                "status": "Pending",
                "output": output_path
            }
            
            # Add to treeview
            self.file_list.insert(
                "",
                "end",
                values=(
                    os.path.basename(path),
                    file_ext[1:].upper(),
                    "Pending",
                    os.path.basename(output_path)
                ),
                tags=(path,)
            )
        
        # Update status
        self.status_var.set(f"{len(self.files)} files ready for conversion")
    
    def convert_all_files(self):
        """Convert all pending files to markdown"""
        # Check if the converter is available
        if not self.converter and not self.check_dependencies():
            return
            
        # Check if there are files to convert
        if not self.files:
            messagebox.showinfo("No Files", "No files to convert. Please add files first.")
            return
        
        # Filter for pending files
        pending_files = {path: info for path, info in self.files.items() 
                         if info["status"] == "Pending"}
        
        if not pending_files:
            messagebox.showinfo("No Pending Files", "All files have already been converted.")
            return
            
        # Start conversion in a separate thread
        self.status_var.set("Converting files...")
        threading.Thread(target=self._convert_files, args=(pending_files,), daemon=True).start()
    
    def _convert_files(self, files_to_convert):
        """Background thread for file conversion"""
        for path, info in files_to_convert.items():
            # Update UI to show conversion is in progress
            self.update_file_status(path, "Converting...")
            
            try:
                # Convert the file with markitdown
                result = self.converter.convert(path)
                
                # Get the markdown content
                try:
                    # Try accessing as markdown property (newer versions)
                    markdown_content = result.markdown
                except AttributeError:
                    # Fall back to text_content for older versions
                    markdown_content = result.text_content
                
                # Ensure the output directory exists
                os.makedirs(os.path.dirname(info["output"]), exist_ok=True)
                
                # Save the markdown to the output file
                with open(info["output"], "w", encoding="utf-8") as f:
                    f.write(markdown_content)
                
                # Update status
                self.update_file_status(path, "Completed")
                self.files[path]["status"] = "Completed"
                
            except Exception as e:
                # Handle errors
                self.update_file_status(path, "Failed")
                self.files[path]["status"] = "Failed"
                print(f"Error converting {path}: {e}")
        
        # Update status bar when done
        completed = sum(1 for info in self.files.values() if info["status"] == "Completed")
        failed = sum(1 for info in self.files.values() if info["status"] == "Failed")
        self.status_var.set(f"Conversion complete: {completed} succeeded, {failed} failed")
        
        # Show message with output directory
        if completed > 0 and self.output_directory:
            messagebox.showinfo(
                "Conversion Complete", 
                f"{completed} files were successfully converted to Markdown.\n\n"
                f"Files were saved to:\n{self.output_directory}"
            )
    
    def update_file_status(self, file_path, status):
        """Update the status of a file in the treeview"""
        # Find the item by tag
        for item in self.file_list.get_children():
            if file_path in self.file_list.item(item, "tags"):
                # Update the status column
                values = self.file_list.item(item, "values")
                new_values = (values[0], values[1], status, values[3])
                self.file_list.item(item, values=new_values)
                break
                
        # Force UI update
        self.update()
    
    def show_context_menu(self, event):
        """Show context menu on right-click"""
        # Get the item that was clicked on
        item = self.file_list.identify_row(event.y)
        if not item:
            return
            
        # Select the item
        self.file_list.selection_set(item)
        
        # Show the context menu
        self.context_menu.post(event.x_root, event.y_root)
    
    def get_selected_file(self):
        """Get the file path of the selected treeview item"""
        selection = self.file_list.selection()
        if not selection:
            return None
            
        # Get the tags (file path) of the selected item
        tags = self.file_list.item(selection[0], "tags")
        if tags:
            return tags[0]
        return None
    
    def open_output(self):
        """Open the output markdown file"""
        file_path = self.get_selected_file()
        if not file_path or self.files[file_path]["status"] != "Completed":
            messagebox.showinfo("Cannot Open", "The file has not been converted successfully.")
            return
            
        output_path = self.files[file_path]["output"]
        if os.path.exists(output_path):
            # Open the file using the default system application
            os.startfile(output_path) if sys.platform == "win32" else os.system(f"open '{output_path}'")
        else:
            messagebox.showinfo("File Not Found", "The output file does not exist.")
    
    def open_output_folder(self):
        """Open the folder containing the output file"""
        file_path = self.get_selected_file()
        if not file_path:
            return
            
        output_path = self.files[file_path]["output"]
        folder_path = os.path.dirname(output_path)
        if os.path.exists(folder_path):
            # Open the folder using the default system application
            os.startfile(folder_path) if sys.platform == "win32" else os.system(f"open '{folder_path}'")
    
    def remove_file(self):
        """Remove the selected file from the list"""
        file_path = self.get_selected_file()
        if not file_path:
            return
            
        # Remove from treeview
        for item in self.file_list.selection():
            self.file_list.delete(item)
            
        # Remove from files dictionary
        if file_path in self.files:
            del self.files[file_path]
            
        # Update status
        self.status_var.set(f"{len(self.files)} files ready for conversion")
    
    def clear_all_files(self):
        """Remove all files from the list"""
        # Clear treeview
        for item in self.file_list.get_children():
            self.file_list.delete(item)
            
        # Clear files dictionary
        self.files.clear()
        
        # Update status
        self.status_var.set("Ready")


if __name__ == "__main__":
    app = MarkdownConverterApp()
    app.mainloop()