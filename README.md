# 📄 PDF & Image Utility GUI

A cross-platform desktop application built with **Python** and **Tkinter** for streamlining document management. This utility allows users to easily merge multiple PDF files and images into a single PDF, or split existing PDFs by page or custom range.

---

## ✨ Features

* **Merge Files:** Combine multiple PDFs and popular image formats (`.jpg`, `.png`) into one consolidated PDF document.
* **PDF Splitting:** Split a selected PDF into separate files (either every page individually or a user-defined page range).
* **Drag-and-Drop:** Seamlessly add files to the list and **reorder files** using drag-and-drop within the listbox.
* **Live Preview:** Instantly view the selected image or navigate through pages of a selected PDF.
* **Theme Support:** Toggleable **Dark/Light Mode** for improved viewing comfort.
* **Dependencies:** Built using standard Python libraries, including `PyPDF2`, `Pillow`, and `PyMuPDF` (`fitz`).

---

## ⚙️ Setup and Installation

### Prerequisites

You must have **Python 3.x** installed on your system.

### Installation Steps

1.  **Save the Code:** Save the entire code block provided into a file named `pdf_utility.py`.
2.  **Install Libraries:** Install the necessary dependencies using `pip`.

    ```bash
    pip install Pillow PyPDF2 PyMuPDF tkinterdnd2
    ```
    *Note: `PyMuPDF` is the package name for the `fitz` library used for PDF rendering.*

---

## 🚀 Usage

1.  **Run the script:**
    ```bash
    python pdf_utility.py
    ```

2.  **Add Files:** Click the **"Browse"** button or drag-and-drop files directly into the list area.
3.  **Manage:**
    * Click on a file to view the **preview**.
    * Click and drag items in the list to **change the merge order**.
    * Select a file and click **"Remove Selected"** to delete it from the list.
4.  **Execute Actions:** Use the buttons at the bottom:
    * **Merge Files:** Combines all items in the list.
    * **Split Selected PDF:** Opens a dialog to split the currently selected PDF.

### Output Location

All merged and split PDF files are automatically saved to your **Desktop** directory.

---

## 📜 License

This project is open-sourced under the **MIT License**. See the `LICENSE` file for details.
