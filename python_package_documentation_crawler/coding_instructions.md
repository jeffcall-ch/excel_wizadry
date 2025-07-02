Specification: Python Documentation Crawler
1. Objective
The goal is to create a Python script that takes a list of package dependencies, finds the official documentation for each package, crawls every page of that documentation, and compiles all the content into a single, comprehensive Markdown file for each package.

2. Prerequisites
The script will require the following third-party Python libraries. Please ensure they are installed.

pip install requests beautifulsoup4 markdownify

3. Step-by-Step Implementation Guide
Step 1: Parse Package Names
Goal: Extract a clean package name from strings like "fastapi (>=0.115.13,<0.116.0)".

Action:
Create a function get_clean_package_name(dependency_string) that takes one of the dependency strings as input and returns only the package name. The function should handle variations like parentheses, version specifiers, and extras (e.g., [cryptography]).

Example Implementation:

import re

def get_clean_package_name(dependency_string: str) -> str:
    """
    Extracts the clean package name from a dependency string.
    e.g., "fastapi (>=0.115.13,<0.116.0)" -> "fastapi"
    e.g., "python-jose[cryptography] (>=3.5.0,<4.0.0)" -> "python-jose"
    """
    # Use a regular expression to find the package name before any special characters
    match = re.match(r"^\s*([a-zA-Z0-9_-]+)", dependency_string)
    if match:
        return match.group(1)
    return "" # Return empty if no match, to be handled later

Step 2: Find the Official Documentation URL from PyPI
Goal: Discover the starting URL for the documentation of a given package.

Action:
Create a function find_docs_url(package_name) that:

Constructs the PyPI URL: https://pypi.org/project/{package_name}/.

Fetches the HTML content of that page.

Uses BeautifulSoup to parse the HTML and find the link to the documentation. This is often labeled "Documentation" or "Homepage". Prioritize the "Documentation" link if available.

Returns the URL as a string. Include error handling for packages not found or links not present.

Example Implementation:

import requests
from bs4 import BeautifulSoup

def find_docs_url(package_name: str) -> str | None:
    """
    Finds the official documentation or homepage URL for a package from its PyPI page.
    """
    pypi_url = f"https://pypi.org/project/{package_name}/"
    print(f"Searching for docs on: {pypi_url}")
    try:
        response = requests.get(pypi_url, timeout=10)
        response.raise_for_status() # Raise an exception for bad status codes
        soup = BeautifulSoup(response.content, 'html.parser')

        # Try to find a link with the text "Documentation" first
        # This selector targets the sidebar project links
        docs_link = soup.select_one('a[href*="documentation"], a[href*="docs"]')
        if docs_link and 'github.com' not in docs_link['href'] and 'gitlab.com' not in docs_link['href']:
             return docs_link['href']

        # If not found, fall back to the "Homepage" link
        homepage_link = soup.select_one('a[aria-label="Homepage"]')
        if homepage_link:
            return homepage_link['href']

        return None
    except requests.RequestException as e:
        print(f"  Error fetching PyPI page for {package_name}: {e}")
        return None

Step 3: Crawl the Documentation Website
Goal: Systematically find every internal link on the documentation website, starting from the main URL.

Action:
Create a function crawl_site(start_url) that:

Initializes two sets: urls_to_visit (seeded with start_url) and visited_urls.

Initializes a list collected_urls to store the final ordered list of pages.

Uses a while loop to run as long as urls_to_visit is not empty.

Inside the loop:

Pops a URL from urls_to_visit.

If the URL is already in visited_urls, continue.

Adds the URL to visited_urls and collected_urls.

Fetches the page content.

Parses the page for all <a> tags.

For each link found:

Use urllib.parse.urljoin to convert relative links (/page.html) to absolute URLs.

Check if the link belongs to the same domain as the start_url.

If it's a valid, internal link and not already visited, add it to urls_to_visit.

Returns the collected_urls list.

Example Implementation:

from urllib.parse import urljoin, urlparse

def crawl_site(start_url: str) -> list[str]:
    """
    Crawls a website starting from a given URL, collecting all internal links.
    """
    if not start_url:
        return []
        
    print(f"Crawling site: {start_url}")
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
                absolute_link = urljoin(current_url, link['href'])
                parsed_link = urlparse(absolute_link)

                # Ensure it's a web link, on the same domain, not an anchor, and not yet visited
                if (parsed_link.scheme in ['http', 'https'] and
                        parsed_link.netloc == base_netloc and
                        '#' not in absolute_link and
                        absolute_link not in visited_urls):
                    urls_to_visit.add(absolute_link)

        except requests.RequestException as e:
            print(f"    Could not fetch {current_url}: {e}")
            continue
            
    return collected_urls

Step 4: Extract and Convert Content
Goal: From a single page URL, extract the main content and convert it to Markdown.

Action:
Create a function get_content_as_markdown(url) that:

Fetches the HTML of the page.

Uses BeautifulSoup to find the main content container. This is the trickiest part. Try a list of common selectors in order, such as main, article, div[role='main'], div#main, div#content, div.content. Use the first one that returns a result. If none match, fall back to the body tag.

Takes the resulting HTML content and converts it to Markdown using the markdownify library.

Gets the page title from the <title> tag to use as a header.

Returns a dictionary containing the title and the Markdown content.

Example Implementation:

from markdownify import markdownify as md

def get_content_as_markdown(url: str) -> dict | None:
    """
    Fetches a URL, extracts its main content, and converts it to Markdown.
    """
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        title = soup.title.string if soup.title else "Untitled Page"
        
        # List of potential main content selectors, from most specific to least
        selectors = ['main', 'article', "div[role='main']", 'div#main', 'div#content', 'div.content']
        main_content = None
        for selector in selectors:
            main_content = soup.select_one(selector)
            if main_content:
                break
        
        if not main_content:
            main_content = soup.body # Fallback

        # Convert the found HTML content to Markdown
        markdown_content = md(str(main_content), heading_style="ATX")
        return {'title': title, 'content': markdown_content}

    except requests.RequestException as e:
        print(f"  Could not process {url}: {e}")
        return None

Step 5: Assemble the Master Script
Goal: Tie all the pieces together into a single, runnable script.

Action:

Define the initial dependencies list.

Loop through each dependency string.

Call get_clean_package_name() to get the name.

Call find_docs_url() to get the documentation starting point.

If a URL is found, call crawl_site() to get all the page URLs.

Iterate through the list of crawled URLs:

Call get_content_as_markdown() for each URL.

Append the resulting title (as a new H2 header: ## Title) and content to a master string for that package.

Save the complete master string to a file named {package_name}.md.

This final orchestration brings all the previous functions together to fulfill the project objective.