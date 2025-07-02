# main_crawler.py
# This script takes a list of Python package dependencies, finds their official
# documentation, crawls the entire documentation website, and compiles the
# content into a single, comprehensive Markdown file for each package.

import re
import requests
import os
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from markdownify import markdownify as md

def get_clean_package_name(dependency_string: str) -> str:
    """
    Extracts the clean package name from a dependency string.
    e.g., "fastapi (>=0.115.13,<0.116.0)" -> "fastapi"
    e.g., "python-jose[cryptography] (>=3.5.0,<4.0.0)" -> "python-jose"
    
    Args:
        dependency_string: The string from the dependency list.

    Returns:
        The cleaned package name.
    """
    # Use a regular expression to find the package name before any special characters
    match = re.match(r"^\s*([a-zA-Z0-9_-]+)", dependency_string)
    if match:
        return match.group(1)
    return "" # Return empty if no match, to be handled later

def find_docs_url(package_name: str) -> str | None:
    """
    Finds the official documentation or homepage URL for a package from its PyPI page.
    
    Args:
        package_name: The clean name of the package.

    Returns:
        The URL to the documentation or None if not found.
    """
    pypi_url = f"https://pypi.org/project/{package_name}/"
    print(f"-> Searching for docs on: {pypi_url}")
    try:
        response = requests.get(pypi_url, timeout=10)
        response.raise_for_status() # Raise an exception for bad status codes
        soup = BeautifulSoup(response.content, 'html.parser')

        # This selector targets sidebar links. We prioritize "Documentation" or "docs".
        # We also try to avoid links pointing directly to code repositories.
        docs_link = soup.find(lambda tag: tag.name == 'a' and ('documentation' in tag.get_text().lower() or 'docs' in tag.get_text().lower()) and 'project-links' in tag.parent.get('class', []))

        if docs_link and docs_link.has_attr('href'):
             print(f"  Found 'Documentation' link: {docs_link['href']}")
             return docs_link['href']

        # If not found, fall back to the "Homepage" link, which is more reliable
        homepage_link = soup.select_one(".sidebar-section a[href*='//']") # A common selector for the homepage
        if homepage_link:
            print(f"  Found 'Homepage' link: {homepage_link['href']}")
            return homepage_link['href']

        print(f"  Could not find a documentation or homepage link for {package_name}.")
        return None
    except requests.RequestException as e:
        print(f"  Error fetching PyPI page for {package_name}: {e}")
        return None

def crawl_site(start_url: str) -> list[str]:
    """
    Crawls a website starting from a given URL, collecting all internal links.
    
    Args:
        start_url: The entry point URL for the documentation website.

    Returns:
        A list of all unique internal URLs found.
    """
    if not start_url:
        return []
        
    print(f"-> Starting crawl of site: {start_url}")
    base_netloc = urlparse(start_url).netloc
    urls_to_visit = {start_url}
    visited_urls = set()
    collected_urls = []

    while urls_to_visit:
        current_url = urls_to_visit.pop()
        if current_url in visited_urls:
            continue

        print(f"  Visiting: {current_url}")
        visited_urls.add(current_url)
        collected_urls.append(current_url)

        try:
            response = requests.get(current_url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')

            for link in soup.find_all('a', href=True):
                absolute_link = urljoin(current_url, link['href']).split('#')[0] # Remove fragments

                # Check if the link is valid and should be visited
                if (urlparse(absolute_link).netloc == base_netloc and
                        absolute_link not in visited_urls and
                        absolute_link not in urls_to_visit):
                    urls_to_visit.add(absolute_link)

        except requests.RequestException as e:
            print(f"    Could not fetch {current_url}: {e}")
            continue
            
    return collected_urls

def get_content_as_markdown(url: str) -> dict | None:
    """
    Fetches a URL, extracts its main content, and converts it to Markdown.
    
    Args:
        url: The URL of the page to process.

    Returns:
        A dictionary with 'title' and 'content' keys, or None on failure.
    """
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        title = soup.title.string.strip() if soup.title else "Untitled Page"
        
        # List of potential main content selectors, from most specific to least
        selectors = ['main', 'article', "div[role='main']", 'div#main', 'div#content', 'div.content']
        main_content = None
        for selector in selectors:
            main_content = soup.select_one(selector)
            if main_content:
                break
        
        # Fallback to the body if no specific content container is found
        if not main_content:
            main_content = soup.body

        if not main_content:
            return None # Can't process if there's no body

        # Convert the found HTML content to clean Markdown
        markdown_content = md(str(main_content), heading_style="ATX", strip=['a'], escape_underscores=False)
        return {'title': title, 'content': markdown_content}

    except requests.RequestException as e:
        print(f"  Could not process {url}: {e}")
        return None

def main():
    """
    Main function to orchestrate the documentation crawling and saving process.
    """
    dependencies = [
        "fastapi (>=0.115.13,<0.116.0)", "uvicorn (>=0.34.3,<0.35.0)",
        "sqlalchemy (>=2.0.41,<3.0.0)", "alembic (>=1.16.2,<2.0.0)"
    ]

    # Create an output directory if it doesn't exist
    output_dir = "documentation_files"
    os.makedirs(output_dir, exist_ok=True)

    for dep in dependencies:
        package_name = get_clean_package_name(dep)
        if not package_name:
            continue
        
        print(f"\n{'='*20}\nProcessing package: {package_name}\n{'='*20}")
        
        docs_url = find_docs_url(package_name)
        if not docs_url:
            print(f"Could not find documentation for {package_name}. Skipping.")
            continue
            
        all_page_urls = crawl_site(docs_url)
        if not all_page_urls:
            print(f"No pages found for {package_name}. Skipping.")
            continue

        full_markdown_content = f"# Full Documentation for {package_name}\n\n"
        
        print(f"-> Processing {len(all_page_urls)} pages...")
        for url in all_page_urls:
            page_data = get_content_as_markdown(url)
            if page_data:
                full_markdown_content += f"\n---\n\n## {page_data['title']}\n\n"
                full_markdown_content += f"*Source URL: {url}*\n\n"
                full_markdown_content += page_data['content']
        
        # Save the final markdown file
        output_filename = os.path.join(output_dir, f"{package_name}.md")
        try:
            with open(output_filename, 'w', encoding='utf-8') as f:
                f.write(full_markdown_content)
            print(f"\nSUCCESS: Successfully saved documentation to {output_filename}")
        except IOError as e:
            print(f"\nERROR: Could not write file {output_filename}: {e}")

if __name__ == "__main__":
    main()