import os
import requests
import msal
from dotenv import load_dotenv
from bs4 import BeautifulSoup
import argparse

load_dotenv()

tenant_id = os.getenv("TENANT_ID")
client_secret = os.getenv("CLIENT_SECRET")
client_id  = os.getenv("CLIENT_ID")
intranet_hr_site_id = os.getenv("MS_SP_ID")

def acquire_token_func():
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    app = msal.ConfidentialClientApplication(
        authority=authority_url,
        client_id=f'{client_id}',
        client_credential=f'{client_secret}'
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token


def find_key_in_json(json_data, target_key):
    if isinstance(json_data, dict):
        for key, value in json_data.items():
            if key == target_key:
                yield value
            elif isinstance(value, (dict, list)):
                yield from find_key_in_json(value, target_key)
    elif isinstance(json_data, list):
        for item in json_data:
            yield from find_key_in_json(item, target_key)

def parse_html(sp_response):
    try:
        text_content = f'Page Description: {sp_response["description"]}\n Content: '
    except KeyError:
        text_content = ""

    for inner_html_content in find_key_in_json(sp_response, 'innerHtml'):
        soup = BeautifulSoup(inner_html_content, 'html.parser')
        text_content += soup.get_text(separator=' ', strip=True)
    return text_content    

def download_file(url, path):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        
        response = requests.get(url)
        
        response.raise_for_status()
        
        with open(path, 'wb') as file:
            file.write(response.content)
        
        print(f"File downloaded successfully and saved to {path}")
    
    except requests.exceptions.RequestException as e:
        print(f"An error occurred while downloading: {e}")
    except OSError as e:
        print(f"An error occurred while creating directories: {e}")

def get_folder_contents(site_id, folder_id, token):
    graph_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{folder_id}/children"
    headers = {
        'Authorization': f'Bearer {token["access_token"]}',
        'Host': 'graph.microsoft.com'
    }

    response = requests.get(graph_endpoint, headers=headers)
    
    if response.status_code == 200:
        return response.json()['value']
    else:
        print(f"Failed to retrieve folder contents: {response.status_code} - {response.text}")
        return []

def is_supported_file(file_name):
    supported_extensions = ('.pdf', '.ppt', '.pptx', '.txt', '.docx', '.rtf')
    return file_name.lower().endswith(supported_extensions)

def download_files_recursive(entries, site_id, token, base_path="data/sharepoint-folders"):
    for page in entries:
        if "@microsoft.graph.downloadUrl" in page:
            if is_supported_file(page['name']):
                file_path = os.path.join(base_path, page['name'])
                download_file(page["@microsoft.graph.downloadUrl"], file_path)
            else:
                print(f"Skipping unsupported file type: {page['name']}")
        elif "folder" in page:
            folder_path = os.path.join(base_path, page['name'])
            print("*"*20)
            print(f"Accessing folder: {folder_path}")
            print("*"*20)
            folder_contents = get_folder_contents(site_id, page['id'], token)
            if folder_contents:
                download_files_recursive(folder_contents, site_id, token, folder_path)
        else:
            print(f"Skipping unrecognized item: {page['name']}")

def scrape_sp_files():
    token  = acquire_token_func()
    headers = {
        'Authorization': f'Bearer {token["access_token"]}',
        'Host': 'graph.microsoft.com'
    }
    
    graph_endpoint = f'https://graph.microsoft.com/v1.0/sites/{intranet_hr_site_id}/drive/root/children'
    response = requests.get(graph_endpoint, headers=headers)
    response.raise_for_status()
    page_info = response.json()
    download_files_recursive(page_info['value'], intranet_hr_site_id, token)

def parse_html(sp_response):
    try:
        text_content = f'Page Description: {sp_response["description"]}\n Content: '
    except KeyError:
        text_content = ""
    for inner_html_content in find_key_in_json(sp_response, 'innerHtml'):
        soup = BeautifulSoup(inner_html_content, 'html.parser')
        text_content += soup.get_text(separator=' ', strip=True)
    return text_content    

def download_html_as_text(url, path, token):
    try:
        print(path)
 
        os.makedirs(os.path.dirname(path), exist_ok=True)

        headers = {
            'Authorization': f'Bearer {token["access_token"]}',
            'Host': 'graph.microsoft.com'
        }
        response = requests.get(url, headers=headers)

        response.raise_for_status()
        
        page_info = response.json()
        title = page_info["title"]
        page_text = parse_html(page_info) 
        file_path = os.path.join(path, f"{title}.txt")
        
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(page_text)
        
        print(f"Page '{title}' downloaded successfully and saved to {file_path}")
    
    except requests.exceptions.RequestException as e:
        print(f"An error occurred while downloading: {e}")
    except OSError as e:
        print(f"An error occurred while creating directories: {e}")

def get_pages(site_id, token):
    graph_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/pages"
    headers = {
        'Authorization': f'Bearer {token["access_token"]}',
        'Host': 'graph.microsoft.com'
    }

    response = requests.get(graph_endpoint, headers=headers)
    
    if response.status_code == 200:
        return response.json()['value']
    else:
        print(f"Failed to retrieve pages: {response.status_code} - {response.text}")
        return []

def download_pages(site_id, token, base_path="data/sharepoint-texts/"):
    pages = get_pages(site_id, token)
    
    for page in pages:
        page_id = page['id']
        print(page["name"])

        page_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/pages/{page_id}/microsoft.graph.sitePage?$expand=canvasLayout"
        
        download_html_as_text(page_url, base_path, token)

def scrape_sp_pages():
    token  = acquire_token_func()
    
    download_pages(intranet_hr_site_id, token)

def find_key_in_json(json_data, target_key):
    if isinstance(json_data, dict):
        for key, value in json_data.items():
            if key == target_key:
                yield value
            elif isinstance(value, (dict, list)):
                yield from find_key_in_json(value, target_key)
    elif isinstance(json_data, list):
        for item in json_data:
            yield from find_key_in_json(item, target_key)

def main():
    parser = argparse.ArgumentParser(description="SharePoint Scraper CLI Tool")
    parser.add_argument("option", choices=["scrape_pages", "scrape_files", "scrape_all"], help="Choose an option to scrape SharePoint pages, files, or both")

    args = parser.parse_args()

    if args.option == "scrape_pages":
        scrape_sp_pages()
    elif args.option == "scrape_files":
        scrape_sp_files()
    elif args.option == "scrape_all":
        scrape_sp_pages()
        scrape_sp_files()

if __name__ == "__main__":
    main()
