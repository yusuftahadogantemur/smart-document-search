# Smart Document Search

A Windows desktop app that searches **inside** your files, not just their names. Remember a word but not which file it was in? Type it, pick a folder, and get every Word, Excel, PowerPoint, PDF, text file and image (via OCR) that contains it.

Built with C# and WinForms on .NET 8.

## Features

- **Searches file contents** of:
  - Word (`.docx`), Excel (`.xlsx`, `.xlsm`), PowerPoint (`.pptx`)
  - PDF (including scanned PDFs through OCR)
  - OpenDocument files (`.odt`, `.ods`, `.odp`)
  - Text and code files (`.txt`, `.csv`, `.json`, `.xml`, `.md`, `.cs`, ...)
  - Images (`.png`, `.jpg`, `.tif`, `.bmp`, `.gif`) through OCR
  - Legacy `.doc`, `.xls`, `.ppt` files on a best-effort basis
- **Also matches file names and folder names**, so a folder called "internship" lists everything inside it.
- **Turkish-aware matching.** Case and accents are ignored (`staj`, `STAJ` and `Staj` are the same; `ç`/`c`, `ş`/`s`, `ı`/`i` are treated alike).
- **Highlighted results** showing the matching sentence, file type badge and folder path.
- **Type filters** (Word, Excel, PDF, Image, ...) with live result counts.
- **Dark and light themes.**
- **Open file** with a double click, or **Show in folder** from the right-click menu.
- **Stop button** that works at any time, plus a 15 second limit per file so one huge file never freezes a search.
- Skips noisy system folders (`Windows`, `Program Files`, `ProgramData`, `AppData`, `node_modules`, `.git`).
- Does not download OneDrive "online-only" files just to read them.

## Download

Prebuilt versions are on the [Releases](../../releases) page. Download the zip, extract the **whole folder**, and run `SmartDocumentSearch.exe`. No .NET installation is needed.

## Build from source

Requirements:

- Windows 10 or 11 (64 bit)
- [.NET 8 SDK](https://dotnet.microsoft.com/download)
- Git

```powershell
git clone https://github.com/yusuftahadogantemur/smart-document-search.git
cd smart-document-search
dotnet run
```

> Use a folder path without Turkish or special characters (for example `C:\Projects\smart-document-search`). The `dotnet` tool can fail to find the project otherwise.

### Enable OCR (optional)

OCR needs the Tesseract language files, which are not stored in this repository. Download them into a `tessdata` folder next to the project file:

```powershell
mkdir tessdata
Invoke-WebRequest https://github.com/tesseract-ocr/tessdata_fast/raw/main/tur.traineddata -OutFile tessdata\tur.traineddata
Invoke-WebRequest https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata -OutFile tessdata\eng.traineddata
```

Without these files the app still works, but images and scanned PDFs are not searched. The status bar will say that OCR did not run.

If OCR fails with a missing DLL error, install the [Visual C++ Redistributable (x64)](https://aka.ms/vs/17/release/vc_redist.x64.exe).

### Create the executable

```powershell
dotnet publish -c Release -r win-x64 --self-contained true -o publish
```

The app is created in the `publish` folder as `SmartDocumentSearch.exe`. Keep all files in that folder together, including `tessdata`.

## How to use

1. Type the word you remember.
2. Choose the folder to scan (the Documents folder is the default).
3. Press **Search**. Results appear while the scan is running.
4. Use the chips above the list to filter by file type.
5. Double-click a result to open it, or right-click it and choose **Show in folder**.

Tip: scan a specific folder such as Documents or Desktop rather than a whole drive. It is much faster and gives cleaner results.

## Notes and limitations
- If the application does not open, extract the file by selecting "Extract All" from the zip file, then try opening the application again. !!!
- Matching is by substring, so `car` also finds `carpet`.
- OCR is slow and can misread characters, so a word inside an image may be missed.
- Password-protected, corrupted or locked files are skipped, and the status bar shows how many.
- Legacy `.doc`, `.xls` and `.ppt` search is approximate.
- Only Windows is supported because the interface uses Windows Forms.

## Built with

- [.NET 8](https://dotnet.microsoft.com/) and Windows Forms
- [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) for Word and PowerPoint
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) for Excel
- [PdfPig](https://github.com/UglyToad/PdfPig) for PDF
- [Tesseract](https://github.com/charlesw/tesseract) for OCR
