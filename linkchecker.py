import os
import re
import requests
from concurrent.futures import ThreadPoolExecutor
from docx import Document

def extract_links_from_docx(docx_path):
    """
    Extract all hyperlinks from a Word document.
    """
    links = set()
    document = Document(docx_path)

    # Loop through paragraphs and extract hyperlinks
    for rel in document.part.rels.values():
        if "hyperlink" in rel.reltype:
            links.add(rel.target_ref)

    # Also look for plaintext links in paragraphs
    link_pattern = re.compile(r'https?://[^\s,;]+')
    for para in document.paragraphs:
        matches = link_pattern.findall(para.text)
        links.update(matches)
    
    return list(links)

def check_link(url):
    """
    Check if a URL is reachable.
    """
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            return (url, 'Alive')
        else:
            return (url, 'Dead (Status: {})'.format(response.status_code))
    except requests.RequestException as e:
        return (url, f'Dead ({str(e)})')

def check_links_parallel(links, max_threads=10):
    """
    Check all links using multi-threading to speed up processing.
    """
    results = []
    with ThreadPoolExecutor(max_threads) as executor:
        futures = [executor.submit(check_link, link) for link in links]
        for future in futures:
            results.append(future.result())
    return results

def main():
    print("Word Document Link Checker")
    docx_path = input("Enter the path to the Word document: ").strip()

    if not os.path.exists(docx_path):
        print("Error: File not found.")
        return

    print("\nExtracting links...")
    links = extract_links_from_docx(docx_path)
    if not links:
        print("No links found in the document.")
        return

    print(f"Found {len(links)} link(s). Checking their status...\n")

    # Check links
    results = check_links_parallel(links)

    # Display results
    for url, status in results:
        color = "\033[92m" if "Alive" in status else "\033[91m"
        print(f"{color}{url} - {status}\033[0m")

    # Save results to a file
    with open("link_check_results.txt", "w") as file:
        for url, status in results:
            file.write(f"{url} - {status}\n")

    print("\nResults saved to 'link_check_results.txt'.")

if __name__ == "__main__":
    main()
