# Little's PDF V2.5📄 - Word, Excel, PPT, JPEG, JPG, PNG to PDF Converter
<img width="256" height="256" alt="image" src="https://github.com/user-attachments/assets/2bff9aef-75bb-489e-8054-71089dfde618" />

Little's PDF is a Quick & clean, efficient, and user-friendly Python-based tool designed to convert Word, Excel, PPT, JPEG, JPG, PNG to PDF. This lets you to handle your daily document conversion and merging needs on Windows without opening heavy apps. This app let you to increase your productivity.


Version 2.5 introduces a massive overhaul inaddition to the Version 1, featuring a brand-new modern GUI built with CustomTkinter, an interactive PDF Preview Studio, advanced page-range extraction, and ZIP packaging capabilities. It completely eliminates the need for heavy, expensive PDF editing software for your daily workflow.

## ✨ What's New in v2.5
* **Modern CustomTkinter UI:** A sleek, borderless, and system-theme-aware interface (Dark/Light mode support).
* **Interactive Preview Studio:** Visually preview PDFs before exporting. Navigate pages and extract specific ranges (e.g., `1, 3-5`).
* **Advanced Export Formats:** Convert PDFs into split individual pages, High-Res JPGs, or Transparent PNGs.
* **ZIP Packaging:** Automatically package large batch conversions or split pages directly into a `.zip` archive.
* **Thread-Safe Processing:** Improved background processing prevents UI freezing and handles file-name collisions smartly (Auto-rename or Overwrite).
* **PDF Compression:** Built-in PyMuPDF compression (garbage collection & deflate) to keep file sizes manageable.

## Screenshots
Just Simply Right click and choose "Send to > Little's PDF tool"

<img width="532" height="156" alt="image" src="https://github.com/user-attachments/assets/6482fa80-1845-4f58-8945-b323c20789ea" />

Single File Processing

<img width="755" height="654" alt="image" src="https://github.com/user-attachments/assets/723bdcf4-bc1f-41f9-9061-1767106091a4" />

Multiple file format Processing

<img width="607" height="536" alt="image" src="https://github.com/user-attachments/assets/135f8372-ca9d-4f52-814d-d414cc179b8d" />

## 🚀 Key Features
### 1. Instant Right-Click Conversion
### 2. The Batch Merger & Processor
### 3. Preview & Export Studio

## 🛠️ Installation Requirements
* **OS:** Windows 10 / 11
* **Dependencies:** Microsoft Office (Word, Excel, PowerPoint) must be installed locally for native COM-based document conversion.

## Quick Setup:
1. Go to the [Releases](#) page and download `Littles_PDF_v2.4_Setup.exe`.
2. Run the installer.
3. The installer automatically registers the context menus and `Send To` shortcuts for you. Start right-clicking immediately!

---

## 💻 Tech Stack
* **Core Logic:** `Python 3`
* **GUI Engine:** `customtkinter`, `tkinter`
* **PDF Manipulation:** `PyMuPDF` (`fitz`), `fpdf`
* **Image Processing:** `Pillow` (`PIL`)
* **Windows Integration:** `pywin32` (`win32com.client`) for silent background MS Office operations.
* [cite_start]**Installer:** `Inno Setup 6` [cite: 1]

## 🤝 Contributing
Feel free to fork this project, submit pull requests, or open an issue if you discover a bug or have a feature request! 

---
*Developed with ☕ and Python by [littlekinz65](https://github.com/littlekinz65/).*

---
Littles PDF v.1
Little's PDF is a fast, efficient, and modern Python-based desktop utility designed to convert Word, Excel, PPT, and major image formats into PDFs directly from the Windows right-click menu.

### Features ###
* **Multi-Format Support**: Effortlessly convert DOCX, XLSX, PPTX, TXT, and major image formats (JPG, PNG, GIF, BMP, TIFF, WEBP) to PDF.
* **Excel Sheet Selector**: Automatically detects multiple sheets in Excel files and allows you to select specific ones for conversion.
* **Drag & Drop Merger**: Easily combine multiple files into a single PDF. Reorder documents via an intuitive GUI before merging.
* **Smart Conflict Handling**: Detects existing files and offers options to overwrite or auto-rename, preventing data loss.
* **Zombie-Free Processes**: Built-in "Zombie Sweeper" logic ensures that hidden Office application processes are properly terminated after use.
* **Modern UI**: Features a custom, compact, and borderless progress bar with real-time status updates.


### Convert files to PDF just by right clicking it ###
<img width="375" height="144" alt="image" src="https://github.com/user-attachments/assets/b8e7eade-9991-4cc3-bce4-e21e89623823" />

### Cute tiny Progress bar that lets you to monitor your status without disturbing you###
<img width="418" height="93" alt="image" src="https://github.com/user-attachments/assets/ec3ddd9f-dda3-4363-86fa-1dd2f7b8e51c" />

### Sheet Selection for Excel Files ###
<img width="354" height="314" alt="image" src="https://github.com/user-attachments/assets/e870100b-1fc8-4459-a26d-42423c17c0f5" />

### Merge Multiple file formats into Single PDF (Merge Excel, word, Images, PPT into Single PDF) ###

Select the files > Right Click > Send to > Merge in Little's PDF. Also Rearrange the Order as per your need

<img width="958" height="488" alt="image" src="https://github.com/user-attachments/assets/6bc2bbee-7609-4c23-84b0-6c4973213f5b" />

### Requirements
* **OS**: Windows
* **Software**: Microsoft Office (Word, Excel, PowerPoint) must be installed for native format conversion.
* **Libraries**:
    * `pywin32`
    * `Pillow`
    * `fpdf`
    * `pypdf`

### How to Use (Note: Ignore all restriction and accept the changes, the code is clean)
1.  **Single Conversion**: Right-click any supported file and choose to "Open with" or run the script via CLI with the file path as an argument.
2.  **Merge**: Pass multiple file paths to the script to trigger the Merger Interface.
3.  **UI Interaction**: Follow the prompts for Excel sheet selection or rearrange files in the merge window using the "Move Up/Down" buttons or direct drag functionality.

### License
* Open to make your changes, Request to keep the content free even if you wish to make any update.
