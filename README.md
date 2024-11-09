# FlipFusion 🔄 - For merging, transforming, and flipping PDFs into new formats!


Welcome to the **PDF Utility App** – where PDFs go to get a serious glow-up! This all-in-one Python application is built to make handling PDFs as smooth as possible. It’s got everything you need and a little more, from converting to Word or Excel, merging files, extracting text, and even setting passwords. If your PDFs had a genie, this app would be it! 🌟

## 🎉 Features

This app packs a punch with features that will transform your PDF workflow:

- **Extract Text**: OCR-based text extraction from scanned PDFs (because PDFs sometimes play hard to get).
- **Merge PDFs**: Combine multiple PDFs into one with a custom file order!
- **Convert to Images**: Convert each PDF page into an image – great for those Instagram-worthy pages. 📸
- **PDF to Word**: Transform your PDFs into editable Word documents.
- **PDF to Excel**: Extract tables straight from PDF to Excel. No more copy-pasting!
- **Set Password**: Secure your PDFs with a password. Don’t let your PDFs roam free!
- **Search Text**: Search for specific text inside your PDFs.

## 🧰 Technologies Used

- **Python** 🐍: The brain behind the app.
- **Tkinter**: For our snazzy GUI.
- **ttkbootstrap**: Dark themes because who doesn’t love a bit of flair?
- **PyMuPDF (Fitz)**: PDF handling and rendering.
- **Pytesseract**: OCR library to read text from images.
- **Pillow**: Image handling.
- **pdf2docx**: Converting PDFs to Word.
- **Tabula-py**: For extracting tables from PDF to Excel.
- **pdf2image**: Converting PDF pages into images.
- **Webbrowser**: To open links to your awesome profiles.

## 🚀 Getting Started

1. **Install the required libraries**: Run this to make sure you have all the essentials:

    ```bash
    pip install tkinter ttkbootstrap pytesseract pymupdf pillow pdf2docx tabula-py pdf2image pandas
    ```

2. **Set up Tesseract**: This app uses Tesseract OCR to read scanned images in PDFs. Install Tesseract [here](https://github.com/tesseract-ocr/tesseract) and update `pytesseract.pytesseract.tesseract_cmd` in the code with your Tesseract installation path. 

3. **Launch the App**: Run the script and explore the endless possibilities:

    ```bash
    python pdf_utility_app.py
    ```

4. **Process PDFs Like a Pro**: The user interface will guide you through each operation. Choose your PDF task, select files, and let the app work its magic! ✨

## 💡 How to Use

### Operation Guide
1. **Extract Text**: Select your PDF and extract readable text.
2. **Merge PDFs**: Choose two or more PDFs, set your preferred order, and merge!
3. **Convert to Images**: Save each page of your PDF as an image.
4. **Convert to Word**: Save your PDF content in a Word document.
5. **Convert to Excel**: Extract tables to an Excel file.
6. **Set Password**: Protect your PDF with a password.
7. **Search Text**: Find specific words in a PDF and get the results instantly.

## 🔒 Note on Security

For password-protecting PDFs, the app will encrypt the document with the password you provide. Just don’t forget your password! Recovery won’t be easy! 🔐

## 🛠 Troubleshooting

- **No text is extracted**: Ensure the PDF is a scanned document for OCR to work properly.
- **Error with paths**: Double-check your Tesseract path in `pytesseract.pytesseract.tesseract_cmd`.
- **Can't find output file**: Check that you've selected an output directory when converting to images or merging files.

## 🎩 Fun Fact

This app was developed in less time than it takes to *actually* read all your PDFs! 😉 Feel free to contribute or share feedback – I’d love to improve this app and make PDF handling even smoother.

---

## 📫 Reach Out

Developed by [Aditya Nalawade](https://www.linkedin.com/in/aditya-nalawade-a4b081297) | [GitHub](https://github.com/Adiiiicodes)
