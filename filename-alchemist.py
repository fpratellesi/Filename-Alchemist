import subprocess
import sys
import os
import re
import json
from datetime import datetime
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

class FileTools:
    FORMAT_OPTIONS = [
        ("alpha2", "Alpha-2 (e.g., US)"),
        ("alpha3", "Alpha-3 (e.g., USA)"),
        ("full_name", "Full Name (e.g., United States)"),
        ("language", "Language (e.g., EN)")
    ]

    CASE_OPTIONS = [
        ("lower", "lowercase"),
        ("title", "Capitalise"),
        ("upper", "UPPERCASE")
    ]

    FORMAT_TOOLTIPS = {
        "alpha2": "Two-letter country code (ISO 3166-1 alpha-2)",
        "alpha3": "Three-letter country code (ISO 3166-1 alpha-3)",
        "full_name": "Complete country name",
        "language": "Primary language code (ISO 639-1)"
    }

    SUPPORTED_EXTENSIONS = {
        '.pdf': 'PDF files',
        '.docx': 'Word documents',
        '.xlsx': 'Excel files (new)',
        '.xls': 'Excel files (legacy)'
 }

    def __init__(self, root):
        self.root = root
        self.root.title("File Alchemist 1.0")
        self.root.geometry("750x820")
        self.root.resizable(False, True)
        
        try:
            self.root.iconbitmap("feather.ico")
        except:
            pass

        # Initialize main variables
        self.current_folder = StringVar()
        self.template_filename = StringVar()
        self.template_parts = []
        self.output_format = StringVar(value=self.FORMAT_OPTIONS[0][1])
        self.case_format = StringVar(value=self.CASE_OPTIONS[0][1])
        self.files_data = []
        self.tooltip = None
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both')
        
        # Initialize tabs
        self.country_code_tab = ttk.Frame(self.notebook)
        self.pdf_remove_tab = ttk.Frame(self.notebook)
        self.pdf_extract_tab = ttk.Frame(self.notebook)
        
        # Add tabs to notebook
        self.notebook.add(self.country_code_tab, text='Country Code Converter')
        self.notebook.add(self.pdf_remove_tab, text='PDF Page Remover')
        self.notebook.add(self.pdf_extract_tab, text='PDF Page Extractor')
        
        # Load country mappings
        self.country_mappings = self.load_country_mappings()
        self.language_to_country = self.create_language_mapping()
        
        # Initialize components for each tab
        self.setup_country_code_converter()
        self.setup_pdf_page_remover()
        self.setup_pdf_page_extractor()

    def create_language_mapping(self):
        language_map = {}
        for country, data in self.country_mappings.items():
            for lang in data['languages'].split('-'):  # Note: changed split separator from ',' to '-'
                if lang not in language_map:
                    language_map[lang] = []
                language_map[lang].append({
                    'country': country,
                    'alpha2': data['alpha2'],
                    'alpha3': data['alpha3']
                })
        return language_map
    
    def load_country_mappings(self):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            json_path = os.path.join(script_dir, 'country_mappings.json')
            print(f"Looking for file at: {json_path}")
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                print(f"Loaded {len(data)} countries")
                return data
        except FileNotFoundError as e:
            print("Error loading file:", e)
            return {"united states": {"alpha2": "us", "alpha3": "usa", "languages": "en"}}

    def setup_country_code_converter(self):
        main_frame = ttk.Frame(self.country_code_tab, padding="10")
        main_frame.grid(row=0, column=0, sticky=(N, W, E, S))
        
        # Folder selection - make entry wider
        ttk.Label(main_frame, text="Folder:").grid(row=0, column=0, sticky=W)
        ttk.Entry(main_frame, textvariable=self.current_folder, width=50).grid(row=0, column=1, sticky=(W, E))  # Changed from 30 to 50
        ttk.Button(main_frame, text="Browse", command=self.browse_folder).grid(row=0, column=2)
        
        # Template input
        template_frame = ttk.LabelFrame(main_frame, text="Template", padding="5")
        template_frame.grid(row=1, column=0, columnspan=3, sticky=(W, E), pady=5)
        
        ttk.Label(template_frame, text="Sample filename:").grid(row=0, column=0, sticky=W)
        ttk.Entry(template_frame, textvariable=self.template_filename, width=50).grid(row=0, column=1, sticky=(W, E))  # Changed from 30 to 50
        ttk.Button(template_frame, text="Analyze", command=self.analyze_template).grid(row=0, column=2)
        
        # Output format selection - make combo wider
        format_frame = ttk.LabelFrame(main_frame, text="Output Format", padding="5")
        format_frame.grid(row=2, column=0, columnspan=3, sticky=(W, E), pady=5)
        
        ttk.Label(format_frame, text="Convert to:").grid(row=0, column=0, sticky=W)
        format_combo = ttk.Combobox(format_frame, 
                                  textvariable=self.output_format,
                                  values=[desc for _, desc in self.FORMAT_OPTIONS],
                                  state="readonly",
                                  width=47)  # Changed from 27 to 47
        format_combo.grid(row=0, column=1, sticky=(W, E), padx=5)
        
        # Case selection - make combo wider
        case_frame = ttk.LabelFrame(main_frame, text="Output Case", padding="5")
        case_frame.grid(row=3, column=0, columnspan=3, sticky=(W, E), pady=5)

        ttk.Label(case_frame, text="Convert case to:").grid(row=0, column=0, sticky=W)
        case_combo = ttk.Combobox(case_frame, 
                                textvariable=self.case_format,
                                values=[desc for _, desc in self.CASE_OPTIONS],
                                state="readonly",
                                width=47)  # Changed from 27 to 47
        case_combo.grid(row=0, column=1, sticky=(W, E), padx=5)
        
        # Space replacement option
        self.replace_spaces = BooleanVar(value=False)
        ttk.Checkbutton(main_frame, text="Replace spaces with underscores", 
                       variable=self.replace_spaces).grid(row=4, column=0, columnspan=3, sticky=W, pady=5)
        
        # Pattern frame - make entries wider
        pattern_frame = ttk.LabelFrame(main_frame, text="Pattern", padding="5")
        pattern_frame.grid(row=5, column=0, columnspan=3, sticky=(W, E), pady=5)
        
        ttk.Label(pattern_frame, text="Search Pattern:").grid(row=0, column=0, sticky=W)
        self.search_pattern = ttk.Entry(pattern_frame, width=50)  # Changed from 30 to 50
        self.search_pattern.grid(row=0, column=1, sticky=(W, E))
        
        ttk.Label(pattern_frame, text="Rename Pattern:").grid(row=1, column=0, sticky=W)
        self.rename_pattern = ttk.Entry(pattern_frame, width=50)  # Changed from 30 to 50
        self.rename_pattern.grid(row=1, column=1, sticky=(W, E))

        # Preview frame with wider treeview
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding="5")
        preview_frame.grid(row=6, column=0, columnspan=3, sticky=(N, W, E, S), pady=5)

        # Add progress display frame
        self.progress_frame = ttk.LabelFrame(main_frame, text="Match Statistics", padding="5")
        self.progress_frame.grid(row=7, column=0, columnspan=3, sticky=(W, E), pady=5)
        
        # Create a frame for the React component
        self.progress_display = ttk.Frame(self.progress_frame, height=50)
        self.progress_display.grid(row=0, column=0, sticky=(W, E))
        
        # Move button frame to row 8
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, columnspan=3, pady=10)
        
        # Set up treeview with column widths
        self.tree = ttk.Treeview(preview_frame, columns=("Original", "New"), show="headings")
        self.tree.heading("Original", text="Original Filename")
        self.tree.heading("New", text="New Filename")
        
        # Set specific column widths
        self.tree.column("Original", width=350)  # Increased column width
        self.tree.column("New", width=350)       # Increased column width
        
        self.tree.grid(row=0, column=0, sticky=(N, W, E, S))
        
        scrollbar = ttk.Scrollbar(preview_frame, orient=VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky=(N, S))
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Status and buttons (these automatically adjust)
        self.status_var = StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=SUNKEN, anchor=W)
        status_bar.grid(row=8, column=0, columnspan=3, sticky=(W, E))
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="Preview", command=self.preview_changes).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Apply Changes", command=self.apply_changes).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Reset", command=self.reset_ui).grid(row=0, column=2, padx=5)

        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        # Show supported file types
        supported_types_str = ", ".join(f"*{ext}" for ext in self.SUPPORTED_EXTENSIONS)
        ttk.Label(main_frame, 
                text=f"Supported file types: {supported_types_str}", 
                font=("", 8, "italic")).grid(row=9, column=0, columnspan=3, sticky=W, pady=(5,0))

         # Add separator line
        ttk.Separator(main_frame, orient='horizontal').grid(row=11, column=0, columnspan=3, sticky=(W, E), pady=10)
        
        # Credits frame
        credits_frame = ttk.Frame(main_frame)
        credits_frame.grid(row=12, column=0, columnspan=3, sticky=(W, E), pady=(0, 10))
        
        # Credits text with styled font
        credits_text = "File Alchemist 1.0 by Federico"
        contact_text = "Please report bugs / ideas for added functions to: federico.pratellesi@oecd.org"
        
        ttk.Label(credits_frame, 
                 text=credits_text,
                 font=("", 9, "bold")).grid(row=0, column=0, columnspan=3)
                 
        ttk.Label(credits_frame, 
                 text=contact_text,
                 font=("", 8, "italic")).grid(row=1, column=0, columnspan=3)
                 
        # Configure the credits frame columns to center the text
        credits_frame.columnconfigure(0, weight=1)        
    
    def split_filename(self, filename):
        print(f"\nDEBUG - split_filename: {filename}")
        filename_lower = filename.lower()
        print(f"DEBUG - filename_lower: {filename_lower}")
        print(f"DEBUG - Available countries:", list(self.country_mappings.keys()))
        matches = []
        
        # Check full country names first
        for country in self.country_mappings:
            pos = filename_lower.find(country.lower())
            print(f"Testing '{country.lower()}' against '{filename_lower}'")
            if pos != -1:
                print(f"DEBUG - Found country match: {country} at position {pos}")
                matches.append({
                    'start': pos,
                    'end': pos + len(country),
                    'value': country,
                    'type': 'country'
                })
        
        # Then check codes
        for country, data in self.country_mappings.items():
            # Check alpha2 code
            code = data['alpha2']
            pos = filename_lower.find(code.lower())
            if pos != -1:
                matches.append({
                    'start': pos,
                    'end': pos + len(code),
                    'value': code,
                    'type': 'alpha2'
                })
            
            # Check alpha3 code
            code = data['alpha3']
            pos = filename_lower.find(code.lower())
            if pos != -1:
                matches.append({
                    'start': pos,
                    'end': pos + len(code),
                    'value': code,
                    'type': 'alpha3'
                })
            
            # Check language codes
            for lang in data['languages'].split(','):
                pos = filename_lower.find(lang.lower())
                if pos != -1:
                    matches.append({
                        'start': pos,
                        'end': pos + len(lang),
                        'value': lang,
                        'type': 'language'
                    })
        
        print(f"DEBUG - Final matches found: {matches}")
        
        if matches:
            type_priority = {'country': 3, 'alpha3': 2, 'alpha2': 1, 'language': 0}
            best_match = max(matches, key=lambda x: (type_priority[x['type']], x['end'] - x['start']))
            
            return [
                filename[:best_match['start']],
                {"type": "code", "value": best_match['value']},
                filename[best_match['end']:]
            ]
        
        return None
    
    def convert_code(self, code, code_type):
        output_format = self.get_format_value()
        country_entry = None
        country_name = None
        
        code_lower = code.lower()
        
        for country, data in self.country_mappings.items():
            if code_type == "country" and country.lower() == code_lower:
                country_entry = data
                country_name = country
                break
            elif code_type == "code":
                if code_lower in [data['alpha2'].lower(), data['alpha3'].lower()] or \
                code_lower in [lang.lower() for lang in data['languages'].split(',')]:
                    country_entry = data
                    country_name = country
                    break
               
        if country_entry:
            if output_format == "alpha2":
                return country_entry['alpha2']
            elif output_format == "alpha3":
                return country_entry['alpha3']
            elif output_format == "full_name":
                return country_name
            elif output_format == "language":
                return country_entry['languages'].split(',')[0]
                
        return None
    
    def analyze_template(self):
            if not self.template_filename.get():
                messagebox.showwarning("Warning", "Please provide a sample filename")
                return
                
            filename = self.template_filename.get()
            # Strip any extension from the template filename
            filename = os.path.splitext(filename)[0]
            print(f"Analyzing filename: {filename}")  # Debug
            
            self.template_parts = self.split_filename(filename)
            print(f"Template parts: {self.template_parts}")  # Debug
            
            if not self.template_parts:
                messagebox.showerror("Error", "Could not identify country/language code in template")
                return
            
            self.generate_pattern()
            print(f"Generated pattern: {self.search_pattern.get()}")  # Debug
            self.preview_changes()
        
    def generate_pattern(self):    
        # Create pattern preserving the exact context from template
        pattern_parts = []  # Corrected indentation
        for part in self.template_parts:
            if isinstance(part, dict) and part['type'] == 'code':
                # Create alternatives for all possible identifiers
                alternatives = []
                # Add full country names
                alternatives.extend(self.country_mappings.keys())
                # Add ISO codes
                for data in self.country_mappings.values():
                    alternatives.extend([data['alpha2'], data['alpha3']])
                    alternatives.extend(data['languages'].split(','))
                
                # Remove duplicates and sort by length (longest first)
                alternatives = sorted(set(alternatives), key=len, reverse=True)
                # Escape special characters and join with OR
                alternatives = [re.escape(alt) for alt in alternatives]
                pattern_parts.append(f'(?P<code>{"|".join(alternatives)})')
            else:
                # Preserve exact formatting/delimiters from template
                pattern_parts.append(re.escape(part))
        
        search_pattern = "^" + "".join(pattern_parts) + "$"
        self.search_pattern.delete(0, END)
        self.search_pattern.insert(0, search_pattern)
        
        # Generate rename pattern preserving template format
        rename_parts = []
        for part in self.template_parts:
            if isinstance(part, dict) and part['type'] == 'code':
                rename_parts.append("{code}")
            else:
                rename_parts.append(part)
        
        rename_pattern = "".join(rename_parts)
        self.rename_pattern.delete(0, END)
        self.rename_pattern.insert(0, rename_pattern)
    
    def get_format_value(self):
        display_text = self.output_format.get()
        for value, text in self.FORMAT_OPTIONS:
            if text == display_text:
                return value
        return "alpha2"

    def preview_changes(self):
        """Preview file renaming changes and display match statistics."""
        if not self.current_folder.get():
            messagebox.showwarning("Warning", "Please select a folder first")
            return

        if not self.search_pattern.get():
            messagebox.showwarning("Warning", "Please provide a sample filename first")
            return

        # Clear existing preview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Initialize counters and lists
        self.files_data = []
        matched_files = 0
        total_files = 0
        unmatched_files = []

        # Compile the pattern
        pattern = self.search_pattern.get()
        try:
            compiled_pattern = re.compile(pattern, re.IGNORECASE)
        except re.error as e:
            messagebox.showerror("Error", f"Invalid pattern: {str(e)}")
            return

        # Process each file in the directory
        for filename in os.listdir(self.current_folder.get()):
            name, ext = os.path.splitext(filename)
            if ext.lower() in self.SUPPORTED_EXTENSIONS:
                total_files += 1
                match = compiled_pattern.match(name)
                if match:
                    code = match.group('code')
                    # First try as country name
                    if code.lower() in [c.lower() for c in self.country_mappings]:
                        new_code = self.convert_code(code, "country")
                    else:
                        # Then try as code
                        new_code = self.convert_code(code, "code")
                    
                    if new_code:
                        # Apply case formatting
                        case_format = self.case_format.get()
                        if case_format == "lowercase":
                            new_code = new_code.lower()
                        elif case_format == "Capitalise":
                            new_code = new_code.title()
                        elif case_format == "UPPERCASE":
                            new_code = new_code.upper()
                        
                        # Generate new filename
                        new_name = self.rename_pattern.get().format(code=new_code)
                        
                        # Handle space replacement
                        if self.replace_spaces.get():
                            new_name = new_name.replace(' ', '_')
                        
                        # Add back the extension
                        new_name_with_ext = new_name + ext
                        
                        # Store and display the match
                        self.files_data.append((filename, new_name_with_ext))
                        matched_files += 1
                        self.tree.insert("", END, values=(filename, new_name_with_ext))
                else:
                    # Track unmatched files
                    unmatched_files.append(filename)

        # Calculate match percentage for the progress bar
        match_percentage = (matched_files / total_files * 100) if total_files > 0 else 0

        # Create HTML for the progress bar
        progress_html = f'''
        <div class="w-full p-4">
            <div class="relative w-full h-8 bg-gray-200 rounded-lg overflow-hidden">
                <div class="absolute top-0 left-0 h-full bg-blue-500 transition-all duration-500" 
                     style="width: {match_percentage}%">
                </div>
                <div class="absolute top-0 left-0 w-full h-full flex items-center justify-center">
                    <span class="text-sm font-medium {('text-white' if match_percentage > 50 else 'text-gray-700')}">
                        {matched_files} / {total_files} files matched ({match_percentage:.1f}%)
                    </span>
                </div>
            </div>
            <div class="mt-2 text-sm text-gray-600">
                {len(unmatched_files)} files did not match the pattern
            </div>
        </div>
        '''

        # Update status text
        status_text = f"Preview: {matched_files}/{total_files} files matched pattern"
        if unmatched_files:
            status_text += "\nUnmatched files:"
            for f in unmatched_files:
                status_text += f"\n - {f}"
        
        self.status_var.set(status_text)

        try:
            # Create and display the progress bar React component
            progress_component = f'''<div id="file-match-progress">
                <FileMatchProgress
                    matched={matched_files}
                    total={total_files}
                    unmatched={len(unmatched_files)}
                />
            </div>'''
            
            # Update the progress display
            self.progress_display.update()
            
        except Exception as e:
            print(f"Error updating progress display: {e}")

        # Show unmatched files popup if any exist
        if unmatched_files:
            messagebox.showinfo(
                "Unmatched Files", 
                f"The following {len(unmatched_files)} files did not match the pattern:\n\n" + 
                "\n".join(f"• {f}" for f in unmatched_files)
            )
       
    def apply_changes(self):
        if not self.files_data:
            messagebox.showwarning("Warning", "Please preview changes first")
            return
            
        errors = []
        success_count = 0
        
        for old_name, new_name in self.files_data:
            try:
                old_path = os.path.join(self.current_folder.get(), old_name)
                new_path = os.path.join(self.current_folder.get(), new_name)
                
                if os.path.exists(new_path):
                    errors.append(f"Skipped {old_name}: Destination file already exists")
                    continue
                    
                os.rename(old_path, new_path)
                success_count += 1
                
            except Exception as e:
                errors.append(f"Error renaming {old_name}: {str(e)}")
                
        message = f"Successfully renamed {success_count} files."
        if errors:
            message += "\n\nErrors:\n" + "\n".join(errors)
            
        messagebox.showinfo("Results", message)
        self.status_var.set(f"Renamed {success_count} files")
        
        self.preview_changes()

    def reset_ui(self):
        self.current_folder.set("")
        self.template_filename.set("")
        self.search_pattern.delete(0, END)
        self.rename_pattern.delete(0, END)
        self.output_format.set(self.FORMAT_OPTIONS[0][1])
        self.case_format.set(self.CASE_OPTIONS[0][1])
        self.replace_spaces.set(False)
        self.files_data = []
        self.template_parts = []
        for item in self.tree.get_children():
            self.tree.delete(item)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.current_folder.set(folder)
            self.status_var.set(f"Selected folder: {folder}")

    def show_tooltip(self, event, text):
        widget = event.widget
        x, y, _, _ = widget.bbox("insert")
        x += widget.winfo_rootx() + 25
        y += widget.winfo_rooty() + 20
        
        self.tooltip = Toplevel(widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        label = ttk.Label(self.tooltip, text=text, background="#ffffe0", relief="solid", borderwidth=1)
        label.pack()

    def hide_tooltip(self, event):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

    def setup_pdf_page_remover(self):
        frame = ttk.Frame(self.pdf_remove_tab, padding="10")
        frame.grid(row=0, column=0, sticky=(N, W, E, S))
        
        # Folder selection
        ttk.Label(frame, text="PDF Folder:").grid(row=0, column=0, sticky=W)
        self.pdf_folder = StringVar()
        ttk.Entry(frame, textvariable=self.pdf_folder, width=40).grid(row=0, column=1, sticky=(W, E))
        ttk.Button(frame, text="Browse", command=self.browse_pdf_folder).grid(row=0, column=2)
        
        # Page numbers input
        ttk.Label(frame, text="Pages to Remove:").grid(row=1, column=0, sticky=W)
        self.pages_to_remove = StringVar()
        ttk.Entry(frame, textvariable=self.pages_to_remove, width=40).grid(row=1, column=1, sticky=(W, E))
        ttk.Label(frame, text="(comma-separated, e.g., 1,2,5)").grid(row=1, column=2, sticky=W)
        
        # Preview and execute buttons
        ttk.Button(frame, text="Preview Changes", command=self.preview_pdf_changes).grid(row=2, column=1)
        ttk.Button(frame, text="Remove Pages", command=self.remove_pdf_pages).grid(row=2, column=2)
        
        # Preview area
        self.pdf_preview = ttk.Treeview(frame, columns=("File", "Pages", "Status"), show="headings")
        self.pdf_preview.heading("File", text="PDF File")
        self.pdf_preview.heading("Pages", text="Pages to Remove")
        self.pdf_preview.heading("Status", text="Status")
        self.pdf_preview.grid(row=3, column=0, columnspan=3, sticky=(N, W, E, S))

        # Configure grid weights
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(3, weight=1)

    def setup_pdf_page_extractor(self):
        frame = ttk.Frame(self.pdf_extract_tab, padding="10")
        frame.grid(row=0, column=0, sticky=(N, W, E, S))
        
        # File selection
        ttk.Label(frame, text="PDF File:").grid(row=0, column=0, sticky=W)
        self.extract_pdf_file = StringVar()
        ttk.Entry(frame, textvariable=self.extract_pdf_file, width=40).grid(row=0, column=1, sticky=(W, E))
        ttk.Button(frame, text="Browse", command=self.browse_extract_pdf).grid(row=0, column=2)
        
        # Page range input
        ttk.Label(frame, text="Start Page:").grid(row=1, column=0, sticky=W)
        self.start_page = StringVar()
        ttk.Entry(frame, textvariable=self.start_page, width=10).grid(row=1, column=1, sticky=W)
        
        ttk.Label(frame, text="End Page:").grid(row=2, column=0, sticky=W)
        self.end_page = StringVar()
        ttk.Entry(frame, textvariable=self.end_page, width=10).grid(row=2, column=1, sticky=W)

        # Output folder selection
        ttk.Label(frame, text="Output Folder:").grid(row=3, column=0, sticky=W)
        self.extract_output_folder = StringVar()
        ttk.Entry(frame, textvariable=self.extract_output_folder, width=40).grid(row=3, column=1, sticky=(W, E))
        ttk.Button(frame, text="Browse", command=self.browse_extract_output).grid(row=3, column=2)
        
        # Extract button
        ttk.Button(frame, text="Extract Pages", command=self.extract_pdf_pages).grid(row=4, column=1)
        
        # Status display
        self.extract_status = StringVar(value="Ready")
        ttk.Label(frame, textvariable=self.extract_status).grid(row=5, column=0, columnspan=3)

        # Configure grid weights
        frame.columnconfigure(1, weight=1)

    def browse_pdf_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.pdf_folder.set(folder)

    def browse_extract_pdf(self):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file:
            self.extract_pdf_file.set(file)
            # Set default output folder to the same as input file
            self.extract_output_folder.set(os.path.dirname(file))

    def browse_extract_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.extract_output_folder.set(folder)

    def preview_pdf_changes(self):
        if not self.pdf_folder.get():
            messagebox.showwarning("Warning", "Please select a PDF folder")
            return
            
        try:
            pages = [int(p.strip()) for p in self.pages_to_remove.get().split(',')]
        except ValueError:
            messagebox.showerror("Error", "Invalid page numbers")
            return
            
        for item in self.pdf_preview.get_children():
            self.pdf_preview.delete(item)
            
        for filename in os.listdir(self.pdf_folder.get()):
            if filename.lower().endswith('.pdf'):
                filepath = os.path.join(self.pdf_folder.get(), filename)
                try:
                    reader = PdfReader(filepath)
                    total_pages = len(reader.pages)
                    status = "OK" if max(pages) <= total_pages else "Invalid pages"
                    self.pdf_preview.insert("", END, values=(filename, self.pages_to_remove.get(), status))
                except Exception as e:
                    self.pdf_preview.insert("", END, values=(filename, "", f"Error: {str(e)}"))

    def remove_pdf_pages(self):
            if not self.pdf_folder.get():
                messagebox.showwarning("Warning", "Please select a PDF folder")
                return
                
            try:
                pages_to_remove = [int(p.strip()) for p in self.pages_to_remove.get().split(',')]
                pages_to_remove = [p - 1 for p in pages_to_remove]  # Convert to 0-based indexing
            except ValueError:
                messagebox.showerror("Error", "Invalid page numbers")
                return
                
            success_count = 0
            error_count = 0
            
            for filename in os.listdir(self.pdf_folder.get()):
                if filename.lower().endswith('.pdf'):
                    file_path = os.path.join(self.pdf_folder.get(), filename)
                    temp_path = os.path.join(self.pdf_folder.get(), f"temp_{filename}")
                    
                    try:
                        # Create a temporary file with the modified content
                        reader = PdfReader(file_path)
                        writer = PdfWriter()
                        
                        for i in range(len(reader.pages)):
                            if i not in pages_to_remove:
                                writer.add_page(reader.pages[i])
                        
                        # Write to temporary file first
                        with open(temp_path, 'wb') as temp_file:
                            writer.write(temp_file)
                        
                        # Close the reader to release the original file
                        reader = None
                        writer = None
                        
                        # Replace original with temporary file
                        os.remove(file_path)
                        os.rename(temp_path, file_path)
                        success_count += 1
                        
                    except Exception as e:
                        error_count += 1
                        if os.path.exists(temp_path):
                            os.remove(temp_path)  # Clean up temp file if it exists
                        
            messagebox.showinfo("Results", 
                            f"Successfully processed {success_count} files\n"
                            f"Errors occurred in {error_count} files")

    def extract_pdf_pages(self):
        if not self.extract_pdf_file.get():
            messagebox.showwarning("Warning", "Please select a PDF file")
            return
            
        if not self.extract_output_folder.get():
            messagebox.showwarning("Warning", "Please select an output folder")
            return
            
        try:
            start_page = int(self.start_page.get())
            end_page = int(self.end_page.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid page numbers")
            return
            
        input_path = self.extract_pdf_file.get()
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(
            self.extract_output_folder.get(), 
            f"{base_name}_excerpt_pages_{start_page}-{end_page}.pdf"
        )
        
        try:
            reader = PdfReader(input_path)
            writer = PdfWriter()
            
            if start_page < 1 or end_page > len(reader.pages):
                messagebox.showerror("Error", "Page numbers out of range")
                return
                
            for i in range(start_page - 1, end_page):
                writer.add_page(reader.pages[i])
                
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
                
            self.extract_status.set(f"Successfully extracted pages {start_page}-{end_page}")
            messagebox.showinfo("Success", "Pages extracted successfully")
            
        except Exception as e:
            self.extract_status.set("Error occurred during extraction")
            messagebox.showerror("Error", str(e))


def main():
    root = Tk()
    app = FileTools(root)
    root.mainloop()


if __name__ == "__main__":
    main()