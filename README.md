# VBA-Tool

A collection of useful VBA macros and utilities for Microsoft Office automation.

## 📋 Table of Contents

- [Overview](#overview)
- [Tools Available](#tools-available)
- [Installation](#installation)
- [Usage](#usage)
- [Features](#features)
- [Contributing](#contributing)
- [License](#license)

## 🎯 Overview

This repository contains a curated collection of VBA (Visual Basic for Applications) tools designed to automate common tasks in Microsoft Office applications. Whether you're working with Word documents, Excel spreadsheets, or PowerPoint presentations, these tools will help you save time and increase productivity.

## 🛠️ Tools Available

### Word Tools

#### Delete Unused Styles
**File:** `DeleteUnusedStyles.vba`

A comprehensive macro to clean up Word documents by removing unused custom styles. This tool helps reduce file size and maintain cleaner document formatting.

**Features:**
- Scans entire document including headers, footers, and tables
- Identifies and removes only custom styles (preserves built-in styles)
- Provides detailed statistics on styles removed
- Includes backup functionality for safety
- Fast and efficient style detection algorithm

**Macros included:**
- `DeleteUnusedStyles()` - Main function to remove unused styles
- `ListAllStyles()` - Display all styles in the document
- `BackupAndDeleteUnusedStyles()` - Create backup before cleaning

## 🚀 Installation

### Method 1: Direct Copy-Paste
1. Open Microsoft Word/Excel/PowerPoint
2. Press `Alt + F11` to open the VBA Editor
3. Go to `Insert > Module`
4. Copy and paste the desired macro code
5. Press `F5` to run or save for later use

### Method 2: Import VBA Files
1. Download the `.vba` file from this repository
2. Open VBA Editor (`Alt + F11`)
3. Right-click on your project in the Project Explorer
4. Select `Import File...`
5. Choose the downloaded `.vba` file

## 📖 Usage

### Delete Unused Styles Tool

#### Quick Start
```vba
' Run this macro to delete unused styles
Sub DeleteUnusedStyles()
```

#### With Backup (Recommended)
```vba
' Run this macro to create backup before deleting
Sub BackupAndDeleteUnusedStyles()
```

#### View All Styles
```vba
' Run this macro to see all styles in document
Sub ListAllStyles()
```

### Step-by-Step Instructions

1. **Open your Word document**
2. **Press `Alt + F11`** to open VBA Editor
3. **Insert > Module** and paste the code
4. **Choose your preferred method:**
   - `DeleteUnusedStyles` - Direct deletion
   - `BackupAndDeleteUnusedStyles` - Creates backup first
   - `ListAllStyles` - View all styles before deciding
5. **Press `F5`** or click Run button
6. **Review the results** in the message box

## ✨ Features

### Safety Features
- **Backup Creation**: Automatically creates timestamped backups
- **Built-in Style Protection**: Only removes custom styles
- **Error Handling**: Robust error handling prevents crashes
- **Progress Indication**: Status bar shows current operation

### Performance Features
- **Efficient Scanning**: Optimized algorithms for large documents
- **Memory Management**: Proper object cleanup
- **Batch Processing**: Handles multiple styles efficiently

### User Experience
- **Clear Feedback**: Detailed statistics and progress messages
- **Multiple Options**: Various ways to run the tool
- **Non-destructive**: Safe operations with backup options

## 🤝 Contributing

We welcome contributions! Here's how you can help:

1. **Fork the repository**
2. **Create a feature branch** (`git checkout -b feature/AmazingFeature`)
3. **Commit your changes** (`git commit -m 'Add some AmazingFeature'`)
4. **Push to the branch** (`git push origin feature/AmazingFeature`)
5. **Open a Pull Request**

### Contribution Guidelines
- Follow consistent code formatting
- Include comments for complex logic
- Test thoroughly before submitting
- Update documentation as needed
- Add examples for new features

## 📝 Code Style

- Use clear, descriptive variable names
- Include comments for complex operations
- Handle errors gracefully
- Follow VBA naming conventions
- Optimize for performance when possible

## 🐛 Issues and Support

If you encounter any issues or have questions:

1. **Check existing issues** in the GitHub repository
2. **Create a new issue** with detailed description
3. **Include your Office version** and operating system
4. **Provide sample files** if relevant (remove sensitive data)

## 🔄 Version History

- **v1.0.0** - Initial release with Delete Unused Styles tool
- More tools coming soon!

## 📋 Requirements

- Microsoft Office 2010 or later
- Windows or Mac OS
- VBA enabled in your Office application

## ⚠️ Disclaimer

These tools are provided as-is. Always backup your documents before running any macro. Test on sample documents first before using on important files.

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🌟 Acknowledgments

- Microsoft Office VBA documentation
- Community contributors and testers
- Stack Overflow VBA community

---

**Made with ❤️ for the VBA community**

*If you find these tools useful, please consider giving this repository a star! ⭐*
