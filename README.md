# xlsx2xliff - Convert Bilingual Excel Files to XLIFF with a GUI

![image](https://github.com/wtigga/xlsx2xliff/assets/7037184/367328c9-f81d-43c9-90f3-2a271ffbb2b1)

**Note:** This project is currently under development. Use it with caution and report any issues you encounter.

## Description

`xlsx2xliff` is a user-friendly Python script that allows you to convert bilingual Excel files into XLIFF (XML Localization Interchange File Format). This tool is especially useful for localization and translation tasks. Key features include:

- Perform Quality Assurance (QA) using xBench.
- Automatically add 'TextID' and 'EXTRA' information to the 'note' section of translatable elements (if such columns are present).
- Runs locally on your computer and does not transfer any data outside of your system.
- Written in Python, and it can be compiled into an executable (*.exe) for use on any computer. It can also be compiled for other platforms such as macOS and Linux from the source code.

## How It Works

1. **Select Source Excel File:** Click the 'Select Excel File' button to open your bilingual Excel file.
2. **Choose Output Location:** Click the 'Save XLIFF to' button to select the location where the XLIFF file will be saved.
3. **Select Source and Target Language Columns:** Use the drop-down menus to choose the source and target language columns from your Excel file.
4. **Perform Conversion:** Click 'Save to XLIFF' to convert the Excel file into XLIFF format.

The program will automatically iterate over all sheets in the Excel file, so ensure that the headers are consistent across all sheets.

## Libraries Used

This project utilizes the following Python libraries:

- [pandas](https://pandas.pydata.org/): For data manipulation and handling Excel files.
- [tkinter](https://docs.python.org/3/library/tkinter.html): For the graphical user interface.
- [xml.etree.ElementTree](https://docs.python.org/3/library/xml.etree.elementtree.html): For working with XML data.
- [webbrowser](https://docs.python.org/3/library/webbrowser.html): For opening URLs in a web browser.

## Prerequisites

- The program is compiled into an executable, so you don't need to install Python separately. However, make sure you have the required Excel files ready for conversion.

## Code

The code for this project is available in the provided Python script file.

## Windows Download
Pre-compiled EXE file can be found under \dist\ directory: dist/excel2xliff.exe

https://github.com/wtigga/xlsx2xliff/blob/e1f9c272d20d524fc168f3dc8da953032f791744/dist/excel2xliff.exe

## License

This project is distributed under the MIT License.

## Acknowledgments

- We would like to thank the open-source community for their valuable contributions.
- This project is maintained by [wtigga](https://github.com/wtigga).

Your feedback and contributions are welcome!

---

**Disclaimer:** Use this tool responsibly and exercise caution when dealing with sensitive data. The authors are not responsible for any loss or damage incurred during the use of this tool.
