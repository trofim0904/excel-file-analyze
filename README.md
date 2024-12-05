
# Excel File Analyze

A **Windows Forms Application** to analyze Excel files, complemented by a simple **PowerShell script** for file size analysis.

## Features

- **Windows Forms Application**:
  - Located in the `src` folder.
  - Provides a graphical interface for analyzing Excel files.
  - Easy-to-use functionality for extracting data insights from Excel files.

- **PowerShell Script**:
  - Found in the `powershell` folder.
  - A lightweight script to analyze and report the size of Excel files.
  - Useful for quick checks on file sizes without opening them.

## Folder Structure

```plaintext
.
├── src
│   └── Windows Forms Application for Excel file analysis
├── powershell
│   └── PowerShell script to analyze Excel file sizes
```

## Getting Started

### Windows Forms Application

1. Navigate to the `src` folder.
2. Open the solution file in Visual Studio.
3. Build and run the application.
4. Use the GUI to load and analyze your Excel files.

### PowerShell Script

1. Go to the `powershell` folder.
2. Open a terminal or PowerShell prompt.
3. Put all excel files to folder with script.
3. Execute the script with:
   ```powershell
   .\get-size.ps1
   ```

## Prerequisites

- For the **Windows Forms Application**:
  - [Microsoft .NET Framework](https://dotnet.microsoft.com/) (or compatible version for building the app).
  - [Visual Studio](https://visualstudio.microsoft.com/) for development.

- For the **PowerShell Script**:
  - PowerShell 5.1 or higher.

## License

This project is licensed under the [MIT License](LICENSE).
