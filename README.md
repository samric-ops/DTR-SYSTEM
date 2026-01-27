# DTR Attendance Log System

A comprehensive web-based application for processing, analyzing, and managing attendance log files from biometric devices. This system reads `.dat` files from ZKTeco, ZKTime, and other biometric devices, providing detailed analytics, hours calculation, and reporting features.

![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=Streamlit&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Plotly](https://img.shields.io/badge/Plotly-3F4F75?style=for-the-badge&logo=plotly&logoColor=white)

## ✨ Live Demo
🔗 **Try it now:** [https://dtr-system.streamlit.app/](https://dtr-system.streamlit.app/)

*Note: Replace with your actual deployment URL*

## 📋 Table of Contents
- [Features](#-features)
- [Supported Formats](#-supported-file-formats)
- [Quick Start](#-quick-start)
- [Installation](#-installation)
- [Usage Guide](#-usage-guide)
- [Deployment](#-deployment)
- [Project Structure](#-project-structure)
- [Customization](#-customization)
- [Troubleshooting](#-troubleshooting)
- [FAQ](#-faq)
- [Support](#-support)
- [License](#-license)

## ✨ Features

### 📤 **File Processing**
- Upload `.dat`, `.txt`, or `.csv` attendance files
- Automatic delimiter detection (comma, tab, space)
- Support for large files (up to 200MB)
- Real-time file validation

### 📊 **Data Analytics**
- Interactive data tables with search and filter
- Records per user analysis
- Daily, weekly, monthly attendance trends
- Hourly distribution charts
- Day-of-week patterns

### ⏱️ **Hours Calculation**
- Automatic IN/OUT record pairing
- Working hours computation
- Overtime calculation
- Break time analysis
- Late arrival detection

### 📈 **Visualization**
- Interactive Plotly charts
- User activity heatmaps
- Attendance rate graphs
- Comparative analysis
- Exportable reports

### 💾 **Export Options**
- Download as CSV
- Export to Excel
- Generate PDF reports
- Save summary as text
- Print-friendly formats

### 🔧 **Advanced Features**
- Multi-user filtering
- Date range selection
- Bulk processing
- Data cleaning tools
- Custom column mapping

## 📁 Supported File Formats

### **Primary Format (ZKTeco/ZKTime)** 
**Example:**
1,2024-01-01 08:00:00,0,1
1,2024-01-01 17:00:00,0,1
2,2024-01-01 08:15:00,0,1

### **Alternative Formats Supported**
1. **Tab-separated:** 
2. **Space-separated:**
3. 1 2024-01-01 08:00:00 0 1
4. **Custom column orders** (configurable in app)

### **Column Explanations**
- **UserID**: Employee/User identification number
- **DateTime**: Timestamp in YYYY-MM-DD HH:MM:SS format
- **Status**: Attendance status (0=Check-in, 1=Check-out, etc.)
- **Verification**: Verification method (1=Fingerprint, 2=RFID, etc.)

## 🚀 Quick Start

### **For End Users (Web Version)**
1. Visit the live app URL
2. Click "Browse" to upload your `.dat` file
3. View and analyze your data immediately
4. No installation required!

### **For Developers (Local Installation)**
```bash
# Clone and run in one command
git clone https://github.com/yourusername/dtr-system.git
cd dtr-system
pip install -r requirements.txt
streamlit run app.py
