# xlsx2xliff
### Convert Bilingual Excel Files to XLIFF with a GUI

![image](https://github.com/wtigga/xlsx2xliff/assets/7037184/e3124d6b-ade6-455d-ba2b-35f245104bae)

## Description

`xlsx2xliff` is a user-friendly Python script that allows you to convert bilingual Excel files into XLIFF (XML Localization Interchange File Format). This tool is especially useful for localization and translation tasks, such as QA with xBench.

Key features:
- Runs locally on your computer and does not transfer any data outside of your system.
- Automatically add 'TextID' and 'EXTRA' information to the 'note' section of translatable elements (if such columns are present).
- Written in Python, and it can be compiled into an executable (*.exe) for use on any computer. It can also be compiled for other platforms such as macOS and Linux from the source code.

## How It Works

1. **Select Source File:** Click the 'Select Excel File' button to open your bilingual Excel file.
2. **Choose Output Location:** Click the 'Save XLIFF to' button to select the location where the XLIFF file will be saved.
3. **Select Source and Target Language Columns:** Use the drop-down menus to choose the source and target language columns from your Excel file.
4. **Select Language Codes**: Use dropdown on the right to pick the corresponding codes.
5. **Perform Conversion:** Click 'Save to XLIFF' to convert the Excel file into XLIFF format.

The program will automatically iterate over all sheets in the Excel file, so ensure that the headers are consistent across all sheets.

## Libraries Used

This project utilizes the following Python libraries:

- [pandas](https://pandas.pydata.org/): For data manipulation and handling Excel files.
- [tkinter](https://docs.python.org/3/library/tkinter.html): For the graphical user interface.
- [xml.etree.ElementTree](https://docs.python.org/3/library/xml.etree.elementtree.html): For working with XML data.
- [webbrowser](https://docs.python.org/3/library/webbrowser.html): For opening URLs in a web browser.

## Prerequisites

- The program is compiled into an executable, so you don't need to install Python separately. However, make sure you have the required Excel files ready for conversion.

## Things to consider

1. Columns that you select as source and target should be identical across all the pages in the source file. There can be many other columns, but those two should have the same name. Otherwise results might not be good.
2. By default, it won't encapsulate inline code (tags) as it is supposed to be done. At the same time, some CAT/TMS won't do that either. You can play with this feature if you expand the window of the script down a little bit, then there will be a field for regex and a checkbox to activate it.


## Windows Download
Pre-compiled EXE file can be found under \dist\ directory as a ZIP archive.


## Acknowledgments

- This project is maintained by [wtigga](https://github.com/wtigga).

---

**Disclaimer:** Use this tool responsibly and exercise caution when dealing with sensitive data. The authors are not responsible for any loss or damage incurred during the use of this tool.
