import os
import datetime
from tkinter import filedialog, messagebox, END, Label, Listbox, Scrollbar
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image, ImageTk
from tkinter import PhotoImage
from tkinter.constants import W, E, LEFT, RIGHT, Y, BOTH, CENTER
from tkinter import CENTER
from tkinter import ttk
from tkinter import Toplevel, simpledialog
from PyPDF2 import PdfMerger, PdfReader, PdfWriter 
from PIL import ImageDraw, ImageFont, ImageGrab
import fitz # PyMuPDF, used for PDF preview.

# --- Global Configuration ---
Output_Dir = os.path.join(os.path.expanduser("~"), "Desktop")
selected_files = []
selected_index = None
preview_label = None
preview_image = None
is_dark_mode = False # Dark mode state variable
drag_start_index = None 

# --- New Global Variables for PDF Paging ---
current_preview_pdf_file = None
current_pdf_page = 0
total_pdf_pages = 0

# --- Core Functions ---

def update_listbox():
    """ Clears and repopulates the file listbox with current selected_files. """
    file_listbox.delete(0, END)
    for i, f in enumerate(selected_files, start=1):
        file_listbox.insert(END, f"{i}. {os.path.basename(f)}")

def clear_preview():
    """ Clears the image from the preview label and resets page controls. """
    global preview_label, preview_image, current_preview_pdf_file
    preview_label.config(image='', text='No Preview')
    preview_image = None
    
    # Hide and reset page controls
    current_preview_pdf_file = None
    page_control_frame.grid_remove() 
    page_label.config(text="")


def browse_files():
    """ Open dialog to select files and add to list. """
    files = filedialog.askopenfilenames(filetypes=[("PDF or Image Files",
                                                    "*.pdf *.jpg *.jpeg *.png")])
    if files:
        for f in files:
            if f not in selected_files:
                selected_files.append(f)
        update_listbox()

def remove_selected():
    """ Removes selected files from the listbox and the selected_files list. """
    global selected_files
    selected_indices = file_listbox.curselection()
    
    if not selected_indices:
        messagebox.showwarning("Warning", "No item selected to remove!")
        return

    # Delete in reverse order to maintain correct indices during deletion.
    for index in reversed(selected_indices):
        del selected_files[index] 
    
    update_listbox()
    clear_preview()

def merge_files():
    """ Merges selected PDFs and images into a single PDF. """
    global selected_files
    if not selected_files:
        messagebox.showerror("Error", "No files selected!")
        return

    timestamp = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    output_filename = f"merged_{timestamp}.pdf"
    output_path = os.path.join(Output_Dir, output_filename)
    
    temp_pdf_paths = []

    try:
        merger = PdfMerger() 
        
        for f in selected_files:
            if f.lower().endswith('.pdf'):
                merger.append(f)
            elif f.lower().endswith(('.png', '.jpg', '.jpeg')):
                try:
                    img = Image.open(f).convert('RGB')
                    temp_pdf_name = f"temp_{os.path.basename(f)}.pdf"
                    temp_pdf_path = os.path.join(Output_Dir, temp_pdf_name)
                    img.save(temp_pdf_path)
                    merger.append(temp_pdf_path)
                    temp_pdf_paths.append(temp_pdf_path)
                except Exception as e:
                    print(f"Error converting image {f} to PDF: {e}")
                    continue

        merger.write(output_path)
        merger.close()

        # Cleanup: remove the temporary PDF files.
        for temp_path in temp_pdf_paths:
            if os.path.exists(temp_path):
                os.remove(temp_path)

        messagebox.showinfo("Success", f"Merging complete!\nExported file: {output_path}")
        
        # Reset application state.
        selected_files.clear()
        update_listbox()
        clear_preview()
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during merging: {e}")
        # Ensure cleanup even on error
        for temp_path in temp_pdf_paths:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass


def split_selected_pdf():
    """ Splits a selected PDF by pages (all or custom range). """
    selected_indices = file_listbox.curselection()

    if not selected_indices:
        messagebox.showerror("Error", "Please select a PDF from the list!")
        return

    if len(selected_indices) > 1:
        messagebox.showwarning("Warning", "Select only ONE PDF to split!")
        return

    file = selected_files[selected_indices[0]]
    
    if not file.lower().endswith(".pdf"):
        messagebox.showerror("Error", "Selected file is not a PDF!")
        return

    try:
        reader = PdfReader(file)
        total_pages = len(reader.pages)
    except Exception as e:
        messagebox.showerror("Error", f"Could not read PDF file: {e}")
        return
    
    split_mode = messagebox.askquestion("Split Mode", 
                                       "Do you want to split ALL pages?\nClick 'Yes' for all pages, 'No' for custom range.")
    
    if split_mode == 'yes':
        # Split every page into its own PDF.
        for i in range(total_pages):
            writer = PdfWriter()
            writer.add_page(reader.pages[i])
            
            # Use original filename in the output
            base_name = os.path.splitext(os.path.basename(file))[0]
            output_filename = f"{base_name}_page_{i+1}.pdf" 
            output_path = os.path.join(Output_Dir, output_filename)
            
            with open(output_path, "wb") as out_file:
                writer.write(out_file)

        messagebox.showinfo("Success", f"Split complete! \n{total_pages} files saved to Desktop.")
    
    else:
        # Custom page range splitting logic
        range_str = simpledialog.askstring("Page Range", 
                                           f"Enter page range (e.g., 1-3) \nTotal pages: {total_pages}")
        if not range_str:
            return

        try:
            start, end = map(int, range_str.split("-"))
            
            if start < 1 or end > total_pages or start > end:
                raise ValueError(f"Range must be between 1 and {total_pages}")

            writer = PdfWriter()
            # Loop from the starting page index (start - 1) up to the ending page index (end).
            for i in range(start - 1, end):
                writer.add_page(reader.pages[i])

            base_name = os.path.splitext(os.path.basename(file))[0]
            output_filename = f"{base_name}_split_range_{start}_to_{end}.pdf"
            output_path = os.path.join(Output_Dir, output_filename)
            
            with open(output_path, "wb") as out_file:
                writer.write(out_file)

            messagebox.showinfo("Success", f"Split complete! \nFile saved to Desktop:\n{output_path}")

        except ValueError as e:
            messagebox.showerror("Error", f"Invalid range format or value: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")


def show_pdf_page(file, page_index):
    """ Renders a specific page of a PDF file to the preview label. """
    global preview_label, preview_image, current_pdf_page, total_pdf_pages
    
    current_pdf_page = page_index
    clear_preview() # Clear previous image, but not page control visibility yet.
    
    try:
        # 1. Load Document and Page
        doc = fitz.open(file)
        total_pdf_pages = len(doc)
        
        if page_index < 0 or page_index >= total_pdf_pages:
            raise IndexError("Page index out of range.")
            
        page = doc.load_page(page_index)
        
        # 2. Render Page to Image
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), colorspace=fitz.csRGB)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close() 

        # 3. Resize Image for Preview
        preview_width, preview_height = 500, 350
        img.thumbnail((preview_width, preview_height), Image.LANCZOS) 

        # 4. Update GUI
        preview_image = ImageTk.PhotoImage(img)
        preview_label.config(image=preview_image, text='')
        
        # 5. Update Page Controls
        page_label.config(text=f"Page {current_pdf_page + 1} of {total_pdf_pages}")
        prev_button.state(['!disabled'] if current_pdf_page > 0 else ['disabled'])
        next_button.state(['!disabled'] if current_pdf_page < total_pdf_pages - 1 else ['disabled'])
        
        # Make sure the controls are visible
        page_control_frame.grid(row=1, column=0, sticky="ew")

    except Exception as e:
        clear_preview()
        page_label.config(text="")
        preview_label.config(text=f"PDF Render Error: {e}")
        doc.close() if 'doc' in locals() and doc else None


def show_preview(index):
    """ Displays a preview of the selected image or PDF page. """
    global preview_label, preview_image, current_preview_pdf_file, current_pdf_page
    
    clear_preview()
    current_pdf_page = 0 # Reset page to 0 for a new file
    
    if not selected_files or index is None or index >= len(selected_files):
        preview_label.config(text="No Preview")
        return
        
    file = selected_files[index]
    
    try:
        # PDF Preview (First Page)
        if file.lower().endswith('.pdf'):
            current_preview_pdf_file = file
            show_pdf_page(file, 0)
            
        # Image Preview
        elif file.lower().endswith(('.png', '.jpg', '.jpeg')):
            current_preview_pdf_file = None
            page_control_frame.grid_remove() # Hide controls for images
            
            img = Image.open(file)
            preview_width, preview_height = 500, 350
            
            # Calculate aspect ratio and resize.
            img.thumbnail((preview_width, preview_height), Image.LANCZOS) 

            preview_image = ImageTk.PhotoImage(img)
            preview_label.config(image=preview_image, text='')
            
        else:
            preview_label.config(text="Unsupported File")
            
    except Exception as e:
        preview_label.config(text=f"Preview Error: {e}")


def next_page():
    """ Moves to the next page of the currently previewed PDF. """
    global current_pdf_page, current_preview_pdf_file, total_pdf_pages
    if current_preview_pdf_file and current_pdf_page < total_pdf_pages - 1:
        current_pdf_page += 1
        show_pdf_page(current_preview_pdf_file, current_pdf_page)

def prev_page():
    """ Moves to the previous page of the currently previewed PDF. """
    global current_pdf_page, current_preview_pdf_file
    if current_preview_pdf_file and current_pdf_page > 0:
        current_pdf_page -= 1
        show_pdf_page(current_preview_pdf_file, current_pdf_page)


def on_click(event):
    """ Handles a mouse click on the listbox (start of selection/drag). """
    global drag_start_index
    
    # Get the index clicked
    try:
        drag_start_index = file_listbox.nearest(event.y)
    except:
        drag_start_index = None
        return

    # Clear all selections and set only the clicked item
    file_listbox.selection_clear(0, END)
    file_listbox.selection_set(drag_start_index)
    
    # Show preview
    show_preview(drag_start_index)

def on_drag(event):
    """ Handles dragging an item for reordering. """
    global drag_start_index, selected_files
    
    if drag_start_index is None:
        return

    new_index = file_listbox.nearest(event.y)
    
    # Check if the drag target is a valid, different index
    if new_index != drag_start_index and 0 <= new_index < len(selected_files):
        
        # Reorder the files in the selected_files list (pop then insert).
        item_to_move = selected_files.pop(drag_start_index)
        selected_files.insert(new_index, item_to_move)

        # Update the listbox display
        update_listbox()
        
        # Keep the selected item highlighted.
        file_listbox.selection_clear(0, END)
        file_listbox.selection_set(new_index)
        
        # Update the start index for the next drag event
        drag_start_index = new_index
        show_preview(new_index)

def on_drop(event):
    """ Handles dropping files from the desktop onto the GUI. """
    global selected_files
    # Splitlist correctly handles file paths, even with spaces
    dropped_files = root.splitlist(event.data)
    
    valid_extensions = ('.pdf', '.png', '.jpg', '.jpeg')
    
    for f in dropped_files:
        # Check if the file has a valid extension and is not already in the list
        if f.lower().endswith(valid_extensions) and f not in selected_files:
            selected_files.append(f)
    
    update_listbox()

def on_select_change(event):
    """ Handler for when the Listbox selection changes. """
    selection = file_listbox.curselection()
    if selection:
        # Check if the selection has actually changed to avoid re-rendering on drag.
        current_selection = selection[0]
        if selected_index != current_selection:
            show_preview(current_selection)
    else:
        clear_preview()

# --- IMPROVED TOGGLE DARK MODE FUNCTION ---
def toggle_dark_mode():
    """ Toggles the application between light and dark themes with better consistency. """
    global is_dark_mode
    is_dark_mode = not is_dark_mode

    if is_dark_mode:
        # Dark Theme Configuration
        
        # 1. Base Window and Ttk Global Background
        root.configure(bg="#2b2b2b")
        style.configure("TFrame", background="#2b2b2b")
        style.configure("TLabel", background="#2b2b2b", foreground="white")
        style.configure("TLabelframe", background="#2b2b2b", foreground="white") # Added
        style.configure("TLabelframe.Label", background="#2b2b2b", foreground="white") # Added

        # 2. Ttk Button Styling
        style.configure("TButton", background="#5c5c5c", foreground="white")
        style.map("TButton", 
                   background=[("active", "#757575"), ("!disabled", "#5c5c5c")],
                   foreground=[("active", "white"), ("!disabled", "white")])
        
        # 3. Standard Tkinter Widgets (Listbox and Preview Label)
        file_listbox.configure(bg="#3c3f41", fg="white", 
                        selectbackground="#007ACC", selectforeground="white",
                        highlightbackground="#5c5c5c", highlightcolor="#007ACC")
        preview_label.configure(bg="#3c3f41", fg="white")
        page_label.configure(bg="#3c3f41", fg="white") # New widget
        
    else:
        # Light Theme Configuration
        
        # 1. Base Window and Ttk Global Background
        root.configure(bg="#f0f0f0")
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", background="#f0f0f0", foreground="black")
        style.configure("TLabelframe", background="#f0f0f0", foreground="black") # Added
        style.configure("TLabelframe.Label", background="#f0f0f0", foreground="black") # Added

        # 2. Ttk Button Styling
        style.configure("TButton", background="#e0e0e0", foreground="black")
        style.map("TButton", 
                   background=[("active", "#cce5ff"), ("!disabled", "#e0e0e0")],
                   foreground=[("active", "black"), ("!disabled", "black")])
        
        # 3. Standard Tkinter Widgets (Listbox and Preview Label)
        file_listbox.configure(bg="white", fg="black",
                        selectbackground="#cce5ff", selectforeground="black",
                        highlightbackground="#cccccc", highlightcolor="#cce5ff")
        preview_label.configure(bg="#f0f0f0", fg="black")
        page_label.configure(bg="#f0f0f0", fg="black") # New widget


# --- GUI Initialization ---

root = TkinterDnD.Tk()
root.title("PDF & Image Merger + Splitter")
root.geometry("1400x800")
root.columnconfigure(1, weight=1)

# Styling using ttk
style = ttk.Style()
style.theme_use("clam")
style.configure("TButton", font=("Segoe UI", 11), padding=6) 
style.configure("TLabel", font=("Segoe UI", 12, "bold"))


# --- Layout: Header ---
header_frame = ttk.Frame(root, padding=10)
header_frame.grid(row=0, column=0, columnspan=2, sticky="ew")
root.columnconfigure(0, weight=1)

ttk.Label(header_frame, text="Select PDFs and Images...").pack(side=LEFT)
ttk.Button(header_frame, text="Browse", command=browse_files).pack(side=LEFT,
                                                                    padx=10)
ttk.Button(header_frame, text="Toggle Dark Mode", 
           command=toggle_dark_mode).pack(side=LEFT, padx=10)


# --- Layout: Middle Section (Listbox & Preview) ---
middle_frame = ttk.Frame(root, padding=10)
middle_frame.grid(row=1, column=0, columnspan=4, sticky="nsew")
middle_frame.columnconfigure(1, weight=1)
root.rowconfigure(1, weight=1) 

# List Frame (Container for the file list)
list_frame = ttk.LabelFrame(middle_frame, text="Selected Files (Drag to Reorder or Drop from Explorer)", padding=10)
list_frame.grid(row=0, column=0, sticky="ns")

# Listbox and Scrollbar
scrollbar = Scrollbar(list_frame)
scrollbar.pack(side=RIGHT, fill=Y)

file_listbox = Listbox(list_frame, width=60, height=20, 
                       yscrollcommand=scrollbar.set, selectmode="extended", 
                       exportselection=False, highlightthickness=1) 
file_listbox.pack(side=LEFT, fill=BOTH, expand=True) 
scrollbar.config(command=file_listbox.yview)

# Listbox Event Bindings
file_listbox.bind("<Button-1>", on_click)
file_listbox.bind("<B1-Motion>", on_drag)
file_listbox.bind("<<ListboxSelect>>", on_select_change)
file_listbox.dnd_bind("<<Drop>>", on_drop) 

# Preview Frame
preview_frame = ttk.LabelFrame(middle_frame, text="Preview", padding=10)
preview_frame.grid(row=0, column=1, padx=20, sticky="nsew")
preview_frame.columnconfigure(0, weight=1)
preview_frame.rowconfigure(0, weight=1)

# The Preview Label (Takes up row 0)
preview_label = Label( 
                        preview_frame,
                        text="No Preview",
                        bg="#f0f0f0", 
                        anchor=CENTER,
                        justify=CENTER,
                        )
preview_label.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

# --- Page Control Frame (Row 1) ---
page_control_frame = ttk.Frame(preview_frame, padding=5)
# Note: It's initially hidden and will be shown by show_pdf_page
page_control_frame.grid(row=1, column=0, sticky="ew") 
page_control_frame.grid_remove() # Start hidden
page_control_frame.columnconfigure(1, weight=1) # Center the label

# Navigation Buttons
prev_button = ttk.Button(page_control_frame, text="< Prev", command=prev_page)
prev_button.grid(row=0, column=0, padx=5)

page_label = Label(page_control_frame, text="", bg="#f0f0f0", fg="black")
page_label.grid(row=0, column=1, sticky="ew")

next_button = ttk.Button(page_control_frame, text="Next >", command=next_page)
next_button.grid(row=0, column=2, padx=5)

# --- Layout: Bottom Buttons ---
bottom_frame = ttk.Frame(root, padding=10)
bottom_frame.grid(row=2, column=0, columnspan=4, sticky="ew")

ttk.Button(bottom_frame, text="Remove Selected",
           command=remove_selected).pack(side=LEFT, padx=10)
ttk.Button(bottom_frame, text="Merge Files",
           command=merge_files).pack(side=LEFT, padx=10)
ttk.Button(bottom_frame, text="Split Selected PDF",
           command=split_selected_pdf).pack(side=LEFT, padx=10)


root.mainloop()