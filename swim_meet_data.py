# Based on v12

# Added rate limiting
# Issue with headless mode where it isn't seeing the tables on the page without the physical page opening
# Issue where individual events are not parsed correctly, but relays are working fine

# Combined Brandon and John's changes

# Import everything needed
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
from urllib.parse import urljoin
import re
import random
from selenium.webdriver.common.by import By
import time

class SwimCloudScraper:
    def __init__(self, delay=1.0, rand_delay_min=8, rand_delay_max=14):
        """
        Initialize the scraper with a delay between requests.

        Args:
            delay: Seconds to wait between requests (default 1.0)
        """
        self.base_url = "https://www.swimcloud.com"
        self.delay = delay
        self.rand_delay_min = rand_delay_min
        self.rand_delay_max = rand_delay_max
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        self.team_name = None

        ## JN- changing selenium chrome to headless
        self._init_selenium()

    def _init_selenium(self):
        """Initialize Selenium with headless Chrome. Disable if you want to see for debugging porpoises"""
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        self.driver = webdriver.Chrome(options=chrome_options)
        print("Initializing headless Chrome for Selenium...")

    def _delay_request(self):
        """Add delay between requests to be respectful to the server."""
        time.sleep(self.delay)

    def find_all_available_sessions(self, url):
        """
        Find all available session links (files with .htm extension) on a meet page.

        Args:
            url: The URL of the meet index page

        Returns:
            list: List of dictionaries containing session info with keys:
                  - 'event_number': Event number (e.g., '1', '3', '112')
                  - 'event_name': Event name (e.g., 'Men 200 Medley Relay')
                  - 'session_type': 'Prelims', 'Finals', 'Swim-off', or None
                  - 'href': The .htm filename
                  - 'full_url': Complete URL to the event page
        """
        self.driver.get(url)
        time.sleep(self.delay)

        # Find all links with .htm extension
        htm_links = self.driver.find_elements(By.XPATH, "//a[contains(@href, '.htm')]")

        sessions = []
        for link in htm_links:
            href = link.get_attribute('href')
            text = link.text.strip()

            # Skip if it's not a valid event link
            if not text or 'Latest Completed Event' in text:
                continue

            # Extract event number from text (e.g., "#1" -> "1")
            event_number = None
            if text.startswith('#'):
                event_number = text.split()[0].replace('#', '')

            # Determine session type
            session_type = None
            text_lower = text.lower()
            if 'prelims' in text_lower:
                session_type = 'Prelims'
            elif 'finals' in text_lower:
                session_type = 'Finals'
            elif 'swim-off' in text_lower or 'swim off' in text_lower:
                session_type = 'Swim-off'

            # Extract just the filename from href
            filename = href.split('/')[-1] if '/' in href else href

            # Build full URL if needed
            full_url = href if href.startswith('http') else f"{url.rsplit('/', 1)[0]}/{filename}"

            sessions.append({
                'event_number': event_number,
                'event_name': text,
                'session_type': session_type,
                'href': filename,
                'full_url': full_url
            })

        print(f"Found {len(sessions)} event sessions")
        return sessions

    def _extract_event_info(self, page_text):
        """
        Extract event number and name from page text.
        Returns: (event_number, event_name, is_relay)
        """
        # Look for pattern like "Event 21  Men 400 Yard Freestyle Relay"
        event_pattern = r'Event\s+(\d+)\s+(.+?)(?:\n|$)'
        match = re.search(event_pattern, page_text)

        if match:
            event_number = match.group(1)
            event_name = match.group(2).strip()
            is_relay = 'Relay' in event_name
            return event_number, event_name, is_relay

        return None, None, False

    def _parse_relay_results(self, page_text, meet_name, meet_url, event_number, event_name):
        """
        Parse relay event results from page text.
        Returns: list of dictionaries with result data
        """
        results = []

        # Split into lines
        lines = page_text.split('\n')

        # Find the start of results (after the header section)
        result_start = 0
        for i, line in enumerate(lines):
            if '==================================================================================' in line:
                result_start = i + 1
                break

        # Parse each result
        i = result_start
        while i < len(lines):
            line = lines[i].strip()

            # Check if this is a result line (starts with rank number)
            rank_match = re.match(r'^\s*(\d+)\s+', line)
            if rank_match:
                # Extract team name and time from this line
                # Pattern: rank, team name, seed time, finals time, points
                parts = re.split(r'\s{2,}', line.strip())

                if len(parts) >= 4:
                    team_name = parts[1]
                    finals_time = parts[3]

                    # Clean up time (remove any letters like 'N', 'A', etc.)
                    finals_time = re.sub(r'[A-Z]', '', finals_time).strip()

                    results.append({
                        'meet_name': meet_name,
                        'meet_url': meet_url,
                        'event_number': event_number,
                        'event_name': event_name,
                        'is_relay': True,
                        'name': team_name,
                        'time': finals_time
                    })

            i += 1

        return results

    def _parse_individual_results(self, page_text, meet_name, meet_url, event_number, event_name):
        """
        Parse individual event results from page text.
        Returns: list of dictionaries with result data
        """
        results = []

        # Split into lines
        lines = page_text.split('\n')

        # Find the start of results
        result_start = 0
        for i, line in enumerate(lines):
            if '==================================================================================' in line:
                result_start = i + 1
                break

        # Parse each result
        i = result_start
        while i < len(lines):
            line = lines[i].strip()

            # Check if this is a result line (starts with rank number)
            rank_match = re.match(r'^\s*(\d+)\s+', line)
            if rank_match:
                # Pattern for individual: rank, name, year, school, seed, finals, points
                # This is more complex and may need adjustment based on actual format
                parts = re.split(r'\s{2,}', line.strip())

                if len(parts) >= 3:
                    # Extract swimmer name (usually second element)
                    swimmer_name = parts[1]

                    # Find the finals time (look for time pattern like 1:23.45)
                    time_pattern = r'\d+:\d+\.\d+|\d+\.\d+'
                    finals_time = None

                    for part in parts:
                        if re.search(time_pattern, part):
                            # This might be the finals time
                            # Usually it's one of the later columns
                            finals_time = re.search(time_pattern, part).group()

                    if finals_time:
                        results.append({
                            'meet_name': meet_name,
                            'meet_url': meet_url,
                            'event_number': event_number,
                            'event_name': event_name,
                            'is_relay': False,
                            'name': swimmer_name,
                            'time': finals_time
                        })

            i += 1

        return results

    def parse_event_page(self, url, meet_name=None, meet_url=None):
        """
        Parse an event results page and extract all relevant data.

        Args:
            url: URL of the event page to parse
            meet_name: Optional meet name (will be extracted if not provided)
            meet_url: Optional meet URL (will use provided URL if not given)

        Returns:
            pandas.DataFrame: DataFrame with columns [meet_name, meet_url, event_number,
                             event_name, is_relay, name, time]
        """
        print(f"Parsing event page: {url}")

        self.driver.get(url)
        time.sleep(self.delay)

        # Get the page text from <pre> tag (results are typically in <pre> tags)
        try:
            pre_element = self.driver.find_element(By.TAG_NAME, 'pre')
            page_text = pre_element.text
        except Exception:
            # Fallback to body text if no <pre> tag
            page_text = self.driver.find_element(By.TAG_NAME, 'body').text

        # Extract meet name if not provided
        if not meet_name:
            meet_name = self._extract_meet_name(page_text)

        # Use the URL as meet_url if not provided
        if not meet_url:
            # Get the base URL (everything before the .htm file)
            meet_url = url.rsplit('/', 1)[0] + '/'

        # Extract event information
        event_number, event_name, is_relay = self._extract_event_info(page_text)

        if not event_number or not event_name:
            print(f"Could not extract event information from {url}")
            return pd.DataFrame()

        print(f"Event {event_number}: {event_name} (Relay: {is_relay})")

        # Parse results based on event type
        if is_relay:
            results = self._parse_relay_results(page_text, meet_name, meet_url,
                                                event_number, event_name)
        else:
            results = self._parse_individual_results(page_text, meet_name, meet_url,
                                                     event_number, event_name)

        print(f"Extracted {len(results)} results")

        # Convert to DataFrame
        df = pd.DataFrame(results)
        return df

    def scrape_entire_meet(self, index_url, output_file='meet_results.xlsx'):
        """
        Scrape all events from a meet and save to Excel.

        Args:
            index_url: URL of the meet index page
            output_file: Path to output Excel file
        """
        print(f"Starting scrape of meet: {index_url}")

        # Get all event sessions
        sessions = self.find_all_available_sessions(index_url)

        # Extract meet name from first page
        if sessions:
            first_event_url = sessions[0]['full_url']
            self.driver.get(first_event_url)
            time.sleep(self.delay)

            try:
                pre_element = self.driver.find_element(By.TAG_NAME, 'pre')
                page_text = pre_element.text
                meet_name = self._extract_meet_name(page_text)
            except:
                meet_name = "Unknown Meet"
        else:
            meet_name = "Unknown Meet"

        print(f"Meet name: {meet_name}")

        # Create or overwrite the Excel file
        all_results = []

        # Parse each event
        for i, session in enumerate(sessions):
            print(f"\nProcessing event {i + 1}/{len(sessions)}: {session['event_name']}")

            try:
                df = self.parse_event_page(session['full_url'],
                                           meet_name=meet_name,
                                           meet_url=index_url)

                if not df.empty:
                    all_results.append(df)

                # Be respectful with delays
                time.sleep(self.delay)

            except Exception as e:
                print(f"Error parsing {session['full_url']}: {e}")
                continue

        # Combine all results
        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)

            # Save to Excel
            print(f"\nSaving {len(final_df)} total results to {output_file}")
            final_df.to_excel(output_file, sheet_name='All Results', index=False)
            print(f"Successfully saved to {output_file}")

            return final_df
        else:
            print("No results found!")
            return pd.DataFrame()

    def close(self):
        """Close the Selenium driver."""
        if hasattr(self, 'driver'):
            self.driver.quit()


# Example usage
if __name__ == "__main__":
    test_mode = True

try:
    # Initialize scraper with 1 second delay between requests
    scraper = SwimCloudScraper(delay=1.0, rand_delay_min=8, rand_delay_max=14)

    # Example URL - replace with your actual meet page URL
    meet_url = "https://swimmeetresults.tech/NCAA-Division-I-Men-2025/evtindex.htm"

    # sessions = scraper.find_all_available_sessions(meet_url)

    # Print results
    # for session in sessions:
    #     print(f"Event #{session['event_number']}: {session['event_name']}")
    #     print(f"  Type: {session['session_type']}")
    #     print(f"  File: {session['href']}")
    #     print(f"  URL: {session['full_url']}")
    #     print()
    # Example: Parse a single event
    event_url = "https://swimmeetresults.tech/NCAA-Division-I-Men-2025/250326lastevt.htm"
    df = scraper.parse_event_page(event_url,
                                  meet_name="2025 NCAA Division I Men's Swimming & Diving",
                                  meet_url="https://swimmeetresults.tech/NCAA-Division-I-Men-2025/")

    print("\nResults preview:")
    print(df.head())

    # Save single event
    df.to_excel('single_event_results.xlsx', sheet_name='Event Results', index=False)

    # Example: Scrape entire meet
    # index_url = "https://swimmeetresults.tech/NCAA-Division-I-Men-2025/index.htm"
    # full_results = scraper.scrape_entire_meet(index_url, output_file='ncaa_meet_results.xlsx')

finally:
    scraper.close()
