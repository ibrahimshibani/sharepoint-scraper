# SharePoint Scraper

This script is designed to retrieve text data from SharePoint SitePages (in `.aspx` format) and all files from a SharePoint drive. It can be used to automate the process of downloading and parsing content from SharePoint, making it easier to manage and analyze your data.

## Features

- **Scrape SharePoint Pages**: Extract text content from SharePoint SitePages.
- **Scrape SharePoint Files**: Download files from a SharePoint drive, supporting various file types such as `.pdf`, `.ppt`, `.pptx`, `.txt`, `.docx`, and `.rtf`.
- **Recursive Download**: Automatically navigate through folders and subfolders to download all supported files.
- **Command-Line Interface (CLI)**: Easily run the script with different options to scrape pages, files, or both.

## Requirements

To run this script, you need to set up the following environment variables:

1. **TENANT_ID**: Your Azure AD tenant ID.
2. **CLIENT_SECRET**: The client secret for your Azure AD application.
3. **CLIENT_ID**: The client ID for your Azure AD application.
4. **MS_SP_ID**: The ID of your SharePoint site.

You can set these environment variables in a `.env` file:

```plaintext
TENANT_ID=your_tenant_id
CLIENT_SECRET=your_client_secret
CLIENT_ID=your_client_id
MS_SP_ID=your_sharepoint_site_id
```

## Usage
Run the script from the command line with one of the following options:
```console
python your_script.py scrape_pages to scrape SharePoint pages.
python your_script.py scrape_files to scrape SharePoint files.
python your_script.py scrape_all to scrape both pages and files.
```
## Example
Here’s an example of how to run the script to scrape all content:
```console

python your_script.py scrape_all
```

The files will be downloaded to a folder called data in the root repository directoy.

## Dependencies
Make sure you have the following Python packages installed:

os
requests
msal
dotenv
beautifulsoup4
argparse
You can install them using pip:
```console
pip install requests msal python-dotenv beautifulsoup4 argparse
```

## License
This project is licensed under the Apache License 2.0. See the LICENSE file for details.

Feel free to customize this further based on your needs! If you have any other questions or need more help, just let me know. 😊

