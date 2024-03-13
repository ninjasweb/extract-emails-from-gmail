## Python Gmail Data Extraction Script

### Overview
This Python script provides a simple yet powerful tool to extract data from Gmail accounts. It leverages the IMAP protocol to connect to a Gmail account, search for specific emails based on criteria such as sender, subject, or date range, and then extracts relevant information from the email bodies.

### Features
- Connects to Gmail accounts securely using IMAP4_SSL.
- Searches for emails based on specified criteria using IMAP search queries.
- Extracts data from email bodies using regular expressions (regex).
- Supports customization for different email formats.
- Saves extracted data to Excel files for further analysis or processing.
- Provides progress bar functionality using tqdm to track processing progress.

### Usage
1. Configure your Gmail username and password as environment variables or use a `.env` file for local development.
2. Customize search criteria and data extraction patterns according to your requirements.
3. Run the script to connect to Gmail, search for emails, extract data, and save it to Excel.

### Requirements
- Python 3.6 or higher
- Required Python packages: `imaplib`, `email`, `openpyxl`, `re`, `tqdm`, `python-dotenv` (optional for `.env` file support)

### Getting Started
To get started, clone this repository and install the required dependencies using pip:

```bash
pip install -r requirements.txt
```

After installing the dependencies, configure your Gmail credentials and customize the script according to your needs. You're now ready to run the script and start extracting data from Gmail!
