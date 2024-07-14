# Societe-Generale



This project provides tools for analyzing and documenting VBA macros in Excel files. It includes scripts to extract data and VBA code from Excel files, convert VBA code into a human-readable format, and a web interface for user interaction.

## Features

- **analyzer.py**: Extracts data from the VBA macros sheet in Excel files.
- **extract_vba.py**: Extracts VBA code from Excel files and converts it into a human-readable format.
- **Flask and HTML UI**: Provides a user-friendly web interface for interacting with the tools.

## Installation

1. **Clone the repository:**
    ```sh
    git clone https://github.com/yourusername/vba-macro-analyzer.git
    cd vba-macro-analyzer
    ```

2. **Set up a virtual environment (optional but recommended):**
    ```sh
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3. **Install the required dependencies:**
    ```sh
    pip install openpyxl pywin32 Flask graphviz nltk
    ```

4. **Install Graphviz:**
    Download and install Graphviz from https://graphviz.org/download/. Make sure to add the Graphviz `bin` directory to your system's PATH.

## Usage

### Analyzer

The `analyzer.py` script extracts data from the VBA macros sheet in an Excel file.

1. **Place your Excel file in the project directory.**
2. **Run the script:**
    ```sh
    python analyzer.py
    ```

### VBA Code Extraction

The `extract_vba.py` script extracts VBA code from an Excel file and converts it into a human-readable format.

1. **Place your Excel file in the project directory.**
2. **Run the script:**
    ```sh
    python extract_vba.py
    ```

### Web Interface

A Flask web application provides a user-friendly interface for interacting with the tools.

1. **Run the Flask application:**
    ```sh
    python app.py
    ```

2. **Open your web browser and navigate to:**
    ```
    http://127.0.0.1:5000
    ```

## Project Structure

vba-macro-analyzer/
│
├── analyzer.py # Script for extracting data from VBA macros sheet
├── extract_vba.py # Script for extracting and converting VBA code
├── app.py # Flask application
├── templates/
│ └── index.html # HTML template for the web interface
├── static/
│ └── styles.css # CSS file for styling the web interface
└── README.md # Project documentation

markdown
