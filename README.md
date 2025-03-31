# Dropbox Paper to Word Document Converter

## Introduction

This notebook provides an automated solution for migrating Dropbox Paper documents to Microsoft Word (.docx) format. It searches your entire Dropbox account for `.paper` files, exports each one as a Markdown file, converts it to a Word document using pandoc, and saves the Word document in a local directory structure that mirrors your Dropbox organization.

## Features

- Automatically finds all Paper documents in your Dropbox account
- Preserves the Dropbox folder structure when saving locally
- Shows progress indicators for long-running operations
- Handles conversion errors gracefully
- Creates a clean Word document for each Paper file

## Prerequisites

Before using this notebook, you'll need:

1. **Dropbox API Key**
   - Create a Dropbox API app at https://www.dropbox.com/developers
   - Under the permissions tab, check all read and write access under individual scopes (**Remember to uncheck all of them once the migration is completed (To ensures data security)**)
   - In the settings tab, generate an access token to use as your API key

2. **Pandoc Installation**
   - Install pandoc from https://pandoc.org/installing.html
   - Pandoc is required for the markdown to Word document conversion

3. **Python Dependencies**
   - dropbox
   - python-dotenv
   - tqdm
   - pypandoc

## Setup Instructions

1. **Configure your environment**
   - Create a `.env` file in the same directory as this notebook
   - Add your Dropbox API key in the format: `paperApiKey='YOUR_API_KEY'`

2. **Install required packages**
   - Run `pip install dropbox python-dotenv tqdm pypandoc` or install from the requirements.txt

3. **Run the notebook**
   - Execute each cell in order
   - The first cell imports dependencies and establishes a connection to Dropbox
   - The second cell defines and runs the conversion function

## How the Conversion Process Works

1. **Discovery**: The notebook recursively searches your entire Dropbox account for files with the `.paper` extension
2. **Export**: Each Paper document is exported to Markdown format using the Dropbox API
3. **Local Storage**: The Markdown file is temporarily saved to a local directory that mirrors its Dropbox path
4. **Conversion**: pypandoc (which uses pandoc) converts the Markdown to a Word document
5. **Cleanup**: The temporary Markdown file is deleted, leaving only the Word document

## Troubleshooting

- **API Connection Issues**: Ensure your API key is correctly set in the .env file and has the appropriate permissions
- **Conversion Errors**: Make sure pandoc is properly installed and accessible from your system PATH
- **File Not Found Errors**: Check that the local directories have proper write permissions
- **Missing Packages**: Install any missing Python packages with pip

## Output Location

Word documents will be saved to a directory structure that mirrors your Dropbox, starting from:
`~/Arca Dropbox/[Your Dropbox Username]/`

---
