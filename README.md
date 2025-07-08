# 📁 AutoFill Report App

A desktop application built with Python and Tkinter that allows users to fill in placeholders in Word (.docx) and Excel (.xlsx) templates with custom values — perfect for generating reports, letters, invoices, or any templated documents.

## 🚀 Features

- Supports .docx and .xlsx templates with {placeholder} syntax
- Dynamic table for adding, editing, and removing placeholder values
- Drag & drop template support
- Save and load placeholder profiles
- Generate filled documents with a single click
- Clear all fields instantly
- Tooltip hints for better usability
- Mouse wheel scrolling in the placeholder table

## 📦 Requirements

- Python 3.8+
- python-docx
- openpyxl
- tkinter (comes with Python)
- tkinterDnD2 (for drag-and-drop support)

Install dependencies:

pip install python-docx openpyxl tkinterdnd2

## 🛠️ How to Use

1. Launch the app:
   python report_app.py

2. Select or drag & drop a .docx or .xlsx template with placeholders like {name}, {date}, etc.

3. Fill in the placeholder values in the table.

4. Click "Generate Output" to create a filled document.

5. Use "Save Profile" and "Load Profile" to reuse placeholder sets.

## 🤝 Contributing

Pull requests are welcome! If you have ideas for new features, feel free to open an issue or fork the repo.

## 📜 License

MIT License — free to use, modify, and distribute.

## 🙌 Acknowledgments

Built by Christopher Hillis  
Powered by Python, Tkinter
