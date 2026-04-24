# 🎓 Student Report Formatter

An automated Python tool that transforms messy academic documents into clean, professional reports in seconds. Designed for students, researchers, and interns, this tool standardizes formatting across project reports, internship manuals, and academic submissions in `.docx` format—with minimal effort.

---

## 🌟 Key Features

### ✨ Automated Styling

* Applies professional fonts such as **Cambria** or **Times New Roman**
* Standardizes font sizes and line spacing for consistency
* Ensures uniform formatting across the entire document

### 📐 Smart Margins & Page Design

* Sets precise **1.0-inch margins** on all sides
* Adds clean, rectangular **page borders** for a polished look

### 📊 Intelligent Table Formatting

* Automatically enhances tables with:

  * Styled header rows
  * Bold, white header text
  * Themed background colors

### 🧠 Heuristic Header Detection

* Detects bold text and upgrades it into structured section headers
* Applies proper sizing (default: **14 pt**) for clear hierarchy

### 📎 Document Merging

* Seamlessly combines multiple Word documents
* Maintains consistent formatting throughout the merged file

### ⚡ Batch Processing

* Formats **all `.docx` files in a folder** in one go
* Saves time when working with multiple reports

---

## 🚀 Getting Started

### ✅ Prerequisites

Make sure you have **Python** installed. Then install the required libraries:

```bash
pip install python-docx pdfminer.six
```

---

## 📦 Installation

1. Clone this repository or download the scripts
2. Place your `.docx` files in the same directory as the script

---

## 🛠️ Usage

### 1. ⚡ Automatic Mode (Recommended)

Formats all Word documents in the folder:

```bash
python manual_converter.py
```

➡️ Output files will be saved with the prefix: `formatted_`

---

### 2. 🎯 Manual Mode

Format a specific file:

```bash
python manual_converter.py "Your_Report_Name.docx"
```

---

### 3. 🔗 Merge Mode

Combine multiple files into a single formatted document:

```bash
python manual_converter.py --merge Final_Report.docx part1.docx part2.docx part3.docx
```

---

## ⚙️ Customization

Modify default settings directly in `manual_converter.py`:

```python
DEFAULT_FONT = "Cambria"
BODY_SIZE = 12
HEADER_SIZE = 14
MARGINS = 1.0
THEME_COLOR = "4D4D80"  # Dark Blue
```

You can tailor the formatting to match your institution’s guidelines.

---

## 📁 Project Structure

```
manual_converter.py     # Core script for formatting and merging
analyze_pdf.py          # Extracts style patterns from PDFs
estimate_margins.py     # Calculates document margin layouts
```

---

## 📄 License

This project is open-source and free to use for **students, educators, and researchers**.

---

## 💡 Why This Tool?

Formatting academic documents manually is tedious and error-prone. This tool eliminates that friction—so you can focus on your content, not formatting.

---

## 🙌 Contributing

Contributions, suggestions, and improvements are welcome! Feel free to fork the project and submit a pull request.

---

## ⭐ Support

If you find this tool useful, consider giving it a star and sharing it with others who might benefit!

---

**Built to make academic life easier—one document at a time.**
