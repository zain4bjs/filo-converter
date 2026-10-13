# PDF-Converter

![<img src="image1.png" width="25"/>](./example%20files/image.png)

If you concern about your PDF files uploading some online website, then here is the nice tool that you can install in your PC or using python files right away convert your PDF files into Word (Docx) format and Images. Supported image formats are 'TIFF', 'JPEG', 'PNG', 'EPS', 'BMP' and 'ICO'. New formats might be added. 

It has 2 functionality for converting PDF to word documents using open source platforms/modules including; pdf2docx, pdf2image, poppler and tesseract. 

It can easily convert PDF files into word documents if the file were converted from word file before. 
Scanned PDF files can also be converted into word documents using OCR (Optical Character Recognition) technology. You just need to choose language of the scanned file. It supports 13 different common languages. 

```
language_formats = ["English", "German", "Turkish",
                    "French", "Spanish", "Italian",
                    "Polish", "Russian", "Ukrainian", "Croatian",
                    "Arabic", "Chinese", "Hindi", "Persian"] # 14 different language format so far supported!
```

Depending on your files, you can choose either option to use.

Video description how to install and use it:

[![Video description how to install and use it](./example%20files/youtubevideo.png)](https://www.youtube.com/watch?v=uhg5xvLCUh8 "Video description how to install and use it")

Python install codes:
```
python -m pip install -r requirements.txt  #install packages
python main.py  #run the program

# you might need to run python3 instead python in mac and linux system.
```

This project is supported by ChatGPT for making the GUI faster.

It is under MIT copyright. You can use it and share with others.

If you find it useful and wants me to improve, please reach out! You can also support me via https://www.buymeacoffee.com/science64

## API (Filo conversion engine)

In this repository's Docker image the folder also runs as an HTTP service
(`api.py`, gunicorn on port 5000, see the top-level README):

- `POST /convert` and `POST /api/convert`: multipart/form-data with the PDF in
  field `file` (optional `mode`: `auto` (default), `editable`, `exact`).
  Success is HTTP 200 with the `.docx` bytes.
- `GET /health`: `{"status": "ok"}`.

Limits: uploads up to `MAX_UPLOAD_MB` (default 30) else 413; PDFs over
`MAX_PDF_PAGES` (default 50) or unreadable by PyMuPDF get 422. A request that
runs past `REQUEST_DEADLINE_SECONDS` (default 95, see the top-level README,
"Request deadline") gets `504 {"error": "Conversion took too long"}`: the
conversion runs in a forked child that is killed with everything it started.
`api.py` imports `../deadline/filo_deadline.py` and `../watermark/`. Errors are
always a short fixed JSON message, e.g. `{"error": "Conversion failed."}`;
exception details are logged server-side only and never returned. The service
sends no CORS headers; add them in a fronting proxy if browsers call it.

The Windows desktop GUI (`python main.py`) looks for Poppler in
`files/poppler-22.04.0/` and Tesseract in `files/Tesseract-OCR/`. Those
Windows builds are not included here: install them and put them on `PATH` (or
in those folders). Linux and macOS use the system `pdftoppm` and `tesseract`.
