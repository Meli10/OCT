# **OCT**

**Version:** 1.0.0  
**Author:** Anthony Meli

## **🔍 Overview**

**OCT** (pronounced OH-CEE-TEE) is a user-friendly application designed to assist users in converting files between several formats, including TSV, CSV, XLSX, and XLS. Additionally, the application ensures that all files are converted to UTF-8 encoding, providing maximum compatibility across different platforms and systems. 🌍

This application is particularly useful for users who work with sensitive data in various formats or simply need a reliable tool to standardize their file encoding and formats quickly.

## **✨ Features**

- **🔄 Multi-Format Conversion:** Convert files between TSV, CSV, XLSX, and XLS formats.
- **🔑 UTF-8 Encoding:** Ensures that all output files are encoded in UTF-8 for broad compatibility.
- **🖥️ User-Friendly Interface:** A clean and simple GUI that guides users through the conversion process step-by-step.
- **📊 Progress Tracking:** Visual progress bar to track the conversion process.
- **⚙️ Cross-Platform Compatibility:** The application is developed using Python and PyQt6, making it compatible with multiple platforms (though the current release is for Windows).

## **💾 Installation**

### **💻 System Requirements**

- **Operating System:** Windows 7 or later
- **Python Version:** The executable is standalone and does not require a separate Python installation.

### **🚀 How to Install**

1. **Download the Application:**
   - [Download the latest release](https://github.com/Meli10/OCT/releases/download/v1.0.0/OCT.exe)
   - Extract the contents of the downloaded zip file to your desired location.

2. **Running the Application:**
   - Simply double-click the `OCT.exe` file to launch the application.
   - No installation is required; the application runs directly from the executable.

## **📚 Usage Instructions**

### **🚦 Launching the Application**

1. Double-click on `OCT.exe` to start the application.
2. The main window titled "Convert Wizard" will appear.

### **🛠️ Steps for File Conversion**

1. **Select Input File:**
   - Click on the "Select Input File" button. 📁
   - Browse and select the file you wish to convert. Supported formats include `.csv`, `.xlsx`, and `.xls`.

2. **Choose Conversion Type:**
   - Select the desired output format: `TSV`, `XLSX`, or `CSV`. 🔄

3. **Select Output Location:**
   - Click on the "Select Output File Location" button. 📁
   - Choose the folder where the converted file should be saved.

4. **Start Conversion:**
   - Click on the "Convert" button. 🚀
   - The progress bar will show the conversion progress. 📊
   - A status message will indicate the success or failure of the conversion process.

### **💻 Command-Line Options (Advanced Users)**

For advanced users who prefer command-line usage, you can build and run the application from the source code with additional customization.

## **⚠️ Known Issues**

- **UTF-8 Encoding Errors:** In rare cases, some files may not convert correctly to UTF-8 encoding. If you encounter this issue, please ensure that your input files are not corrupted and are in a supported format.

## **🤝 Contributing**

If you would like to contribute to the development of this application, please feel free to submit a pull request or report issues. Contributions are always welcome!

## **📞 Support**

For support, questions, or issues, please contact the owner of this repository or [report an issue](#).
