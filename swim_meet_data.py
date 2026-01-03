# Based on v12 This version is 1.0.0.2 (added individual event parsing)

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


class SwimMeetScraper:
    def __init__(self, delay=1.0, rand_delay_min=8, rand_delay_max=14, headless=False):
        """
        Initialize the scraper with a delay between requests.
        """

        self.delay = delay
        self.rand_delay_min = rand_delay_min
        self.rand_delay_max = rand_delay_max
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        self.team_name = None
        self.headless = headless

        self._init_selenium(headless=headless)

    def _init_selenium(self, headless):
        chrome_options = Options()
        if headless:
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)')
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_argument('--disable-extensions')
            chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
            print("Initializing the headless horseChrome. ..")
        else:
            print("Initializing Chrome *with head*...")

        self.driver = webdriver.Chrome(options=chrome_options)

    # ---------------- DIVING PARSER (FIXED, EVERYTHING ELSE UNCHANGED) ---------------- #

    def _parse_diving_results(self, page_text, meet_name, meet_url, event_number, event_name):
        results = []
        lines = page_text.split('\n')

        result_start = 0
        for i, line in enumerate(lines):
            if 'Preliminaries' in line:
                result_start = i + 1
                break

        YEAR_TOKENS = {'FR', 'SO', 'JR', 'SR', '5Y'}

        i = result_start
        while i < len(lines):
            line = lines[i].strip()

            if not line or line.startswith('=='):
                i += 1
                continue

            rank_match = re.match(r'^(\d+)\s+', line)
            if not rank_match:
                i += 1
                continue

            parts = line.split()
            rank = parts[0]

            name_parts = []
            year = None
            school = None

            name_start = None
            for idx in range(1, len(parts)):
                if ',' in parts[idx]:
                    name_start = idx
                    break

            if name_start is not None:
                for j in range(name_start, len(parts)):
                    if parts[j] in YEAR_TOKENS:
                        year = parts[j]
                        school = ' '.join(parts[j + 1:-1]) if j + 1 < len(parts) else None
                        break
                    name_parts.append(parts[j])

            name = ' '.join(name_parts) if name_parts else None

            score = None
            for part in reversed(parts):
                if re.match(r'^\d+\.\d+$', part):
                    score = part
                    break

            if rank != "2025" and name:
                results.append({
                    'meet_name': meet_name,
                    'meet_url': meet_url,
                    'event_number': event_number,
                    'event_name': event_name,
                    'Rank': rank,
                    'Name': name,
                    'Year': year,
                    'School': school,
                    'Score': score
                })

            i += 1

        return results

    def find_all_available_sessions(self, url):
        """
        Find all available session links (files with .htm extension) on a meet page.

        Args:
            url: The URL of the meet index page

        Returns:
            list: List of dictionaries containing session info
        """
        self.driver.get(url)
        time.sleep(self.delay)

        # Not finding .htm links properly, testing stuff
        # Debugging - this works!! Need to switch to frame first
        # "It's working!" --Anakin, sometime
        frame = self.driver.find_element(By.TAG_NAME, 'frame')
        self.driver.switch_to.frame(frame)
        htm_links = self.driver.find_elements(By.XPATH, "//a[contains(@href, '.htm')]")
        print(f"DEBUG: Found {len(htm_links)} .htm links inside frame")

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
            else:
                session_type = 'Relay Only'

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
        self.driver.switch_to.default_content()
        return sessions

    def _extract_meet_name(self, page_text):
        """Extract meet name from page text."""
        lines = page_text.strip().split('\n')
        for i, line in enumerate(lines):
            if 'Championship' in line or 'Meet' in line:
                # Often the meet name is in the first few lines
                return line.strip()
        return "Unknown Meet"

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

    def _determine_relay_distances(self, event_name):
        """
        Determine the distances for each leg based on relay type.
        Returns list of distances (e.g., [50, 100, 150, 200] for 200 relay)
        """
        if '200' in event_name and 'Relay' in event_name:
            return [50, 100, 150, 200]
        elif '400' in event_name and 'Relay' in event_name:
            return [100, 200, 300, 400]
        elif '800' in event_name and 'Relay' in event_name:
            return [200, 400, 600, 800]
        else:
            # Default to 4 legs with unknown distances
            return [1, 2, 3, 4]

    def _parse_relay_results(self, page_text, meet_name, meet_url, event_number, event_name):
        """
        Parse relay event results from page text with individual swimmer splits.
        Returns: list of dictionaries with detailed split data
        """
        results = []

        # Split into lines
        lines = page_text.split('\n')

        # Find the start of results
        # The number of ==='s changes by event, and happens twice- changed logic to find any === and take second one
        result_start = 0
        second_separator_found = 0
        for i, line in enumerate(lines):
            if line.startswith('========='):
                if second_separator_found == 1:
                    result_start = i + 1
                    break
                second_separator_found += 1

        # Parse each team result
        i = result_start
        while i < len(lines):
            line = lines[i].strip()

            # Stop if we've reached the team rankings or end
            if 'Team Rankings' in line or line.startswith('Men -') or line.startswith('Women -'):
                print(f"DEBUG: Reached end section at line {i}")
                break

            # Check if this is a result line (starts with rank number)
            rank_match = re.match(r'^\s*(\d+)\s+', line)
            if rank_match:
                team_name = None
                swimmers = []  # List of (order, name) tuples
                splits_lines = []

                # Extract team name from the first line
                # Pattern: "1 Tennessee    2:42.41    2:42.30N  40"
                # Team name is between rank and the first time (contains colon or multiple digits)
                parts = line.split()
                if len(parts) >= 2:
                    # Find where the times start (they contain colons or are seed/finals times)
                    team_parts = []
                    for j, part in enumerate(parts[1:], 1):
                        # Check if this looks like a time (has colon or is a number with decimal and letters)
                        if ':' in part or re.match(r'^\d+\.\d+[A-Z]*$', part):
                            # Everything before this is the team name
                            team_name = ' '.join(parts[1:j])
                            break

                    # If we didn't find a time pattern, team name might be just the second element
                    if not team_name and len(parts) > 1:
                        team_name = parts[1]

                # Move to next line to start looking for swimmers
                i += 1

                # Next lines contain swimmer names
                # Pattern: "1) Caribe, Guilherme JR          2) r:0.23 Taylor, Lamar 5Y"
                while i < len(lines):
                    next_line = lines[i].strip()
                    # Swimmer lines start with numbers followed by )
                    if re.match(r'^\d+\)', next_line):
                        # Parse all swimmers in this line
                        # Split by pattern of digit followed by )
                        parts = re.split(r'(?=\d+\))', next_line)

                        for part in parts:
                            part = part.strip()
                            if not part:
                                continue

                            # Extract order number and name
                            # Pattern: "1) Caribe, Guilherme JR" or "2) r:0.23 Taylor, Lamar 5Y"
                            match = re.match(r'(\d+)\)\s*(?:r:[\d.+-]+\s*)?(.+)', part)
                            if match:
                                order = int(match.group(1))
                                name = match.group(2).strip()
                                swimmers.append((order, name))
                        i += 1
                    else:
                        break

                # Find the splits lines (contains the actual split times)
                # Lines start with "r:" or just have times
                while i < len(lines):
                    next_line = lines[i].strip()
                    if next_line.startswith('r:') or re.match(r'^\d+\.\d+', next_line):
                        splits_lines.append(next_line)
                        i += 1
                        # Continue reading lines that look like splits
                        while i < len(lines):
                            cont_line = lines[i].strip()
                            # Check if this is a continuation of splits (has time patterns)
                            if re.search(r'\d+:\d+\.\d+|\d+\.\d+', cont_line) and not re.match(r'^\d+\s+\S', cont_line):
                                splits_lines.append(cont_line)
                                i += 1
                            else:
                                break
                        break
                    i += 1

                # Combine all split lines into one string
                splits_text = ' '.join(splits_lines)

                # Parse splits from the combined splits text
                if splits_text and swimmers:
                    # Extract all time values
                    time_pattern = r'(\d+:\d+\.\d+|\d+\.\d+)'
                    all_times = re.findall(time_pattern, splits_text)

                    leg_data = []
                    times_idx = 0

                    # Skip reaction time if present (r:+0.58 becomes 0.58)
                    if times_idx < len(all_times) and float(all_times[times_idx]) < 1.0:
                        times_idx += 1

                    # First leg: split, leg, cumulative (cumulative appears twice)
                    if times_idx + 2 < len(all_times):
                        split_time = all_times[times_idx]  # 19.28
                        leg_time = all_times[times_idx + 1]  # 40.57
                        cumulative = all_times[times_idx + 2]  # 40.57 (duplicate)
                        leg_data.append((split_time, leg_time, cumulative))
                        times_idx += 3

                    # Remaining legs follow pattern:
                    # intermediate_cumulative, split, cumulative, leg
                    # We want: split, leg, cumulative
                    while times_idx < len(all_times) and len(leg_data) < len(swimmers):
                        # Skip intermediate cumulative (e.g., 59.68)
                        times_idx += 1

                        if times_idx >= len(all_times):
                            break

                        # Get split time (e.g., 19.11)
                        split_time = all_times[times_idx]
                        times_idx += 1

                        if times_idx >= len(all_times):
                            break

                        # Get cumulative time (e.g., 1:21.59)
                        cumulative = all_times[times_idx]
                        times_idx += 1

                        if times_idx >= len(all_times):
                            break

                        # Get leg time (e.g., 41.02)
                        leg_time = all_times[times_idx]
                        times_idx += 1

                        leg_data.append((split_time, leg_time, cumulative))

                    # Create a result entry for each swimmer
                    for idx, (order, name) in enumerate(swimmers):
                        if idx < len(leg_data):
                            split_time, leg_time, cumulative = leg_data[idx]

                            results.append({
                                'meet_name': meet_name,
                                'meet_url': meet_url,
                                'event_number': event_number,
                                'event_name': event_name,
                                'is_relay': True,
                                'Team Name': team_name,
                                'Name': name,
                                'Order': order,
                                'Split': split_time,
                                'Leg': leg_time,
                                'Cumulative': cumulative
                            })
            else:
                i += 1

        return results

    def _parse_individual_results(self, page_text, meet_name, meet_url, event_number, event_name):
        """
        Parse individual event results from page text with all splits.
        Returns: list of dictionaries with detailed split data (up to 33 splits)
        """
        results = []

        # Split into lines
        lines = page_text.split('\n')

        # Find the start of results
        result_start = 0
        second_separator_found = 0
        for i, line in enumerate(lines):
            if line.startswith('========='):
                if second_separator_found == 1:
                    result_start = i + 1
                    break
                second_separator_found += 1

        print(f"DEBUG: Starting individual results parse at line {result_start} of {len(lines)}")

        # Parse each result
        i = result_start
        results_count = 0
        while i < len(lines):
            line = lines[i].strip()

            # Check if this is a result line (starts with rank number)
            rank_match = re.match(r'^\s*(\d+)\s+', line)
            if rank_match:
                rank = rank_match.group(1)
                results_count += 1
                if results_count % 10 == 0:
                    print(f"DEBUG: Processed {results_count} swimmers...")

                parts = line.split()

                # Extract swimmer name and year/school
                swimmer_name = None
                year = None
                school = None

                # Find name (contains comma)
                for j in range(1, len(parts)):
                    if ',' in parts[j]:
                        swimmer_name = parts[j]
                        # Check if next parts are part of name or year/school
                        if j + 1 < len(parts):
                            # Year is typically 2-3 characters (JR, SR, SO, FR, 5Y)
                            if len(parts[j + 1]) <= 3 and parts[j + 1].isupper():
                                year = parts[j + 1]
                                if j + 2 < len(parts):
                                    school = parts[j + 2]
                            else:
                                # Part of name
                                swimmer_name += ' ' + parts[j + 1]
                                if j + 2 < len(parts) and len(parts[j + 2]) <= 3:
                                    year = parts[j + 2]
                                    if j + 3 < len(parts):
                                        school = parts[j + 3]
                        break

                # Find the finals time (last time in the line)
                time_pattern = r'\d+:\d+\.\d+|\d+\.\d+'
                finals_time = None
                for part in reversed(parts):
                    if re.match(time_pattern, part):
                        finals_time = re.sub(r'[A-Z]', '', part)
                        break

                # Now collect all split lines for this swimmer
                i += 1
                splits_lines = []
                splits_line_count = 0
                max_splits_lines = 20  # Safety limit to prevent infinite loops

                # Continue reading lines until we hit another result or separator
                while i < len(lines) and splits_line_count < max_splits_lines:
                    next_line = lines[i].strip()

                    # Stop if we hit another result line or empty line or section divider
                    if re.match(r'^\s*(\d+)\s+', next_line) or next_line.startswith(
                            '--') or next_line == '' or next_line.startswith('===') or 'Team Rankings' in next_line:
                        break

                    # Check if this line contains splits (has time patterns)
                    if re.search(r'\d+:\d+\.\d+|\d+\.\d+', next_line):
                        splits_lines.append(next_line)
                        splits_line_count += 1
                        i += 1
                    else:
                        i += 1
                        break

                # Parse all splits from the collected lines
                splits_data = {}
                if splits_lines:
                    all_splits_text = ' '.join(splits_lines)

                    # Extract all times for processing
                    time_pattern = r'(\d+:\d+\.\d+|\d+\.\d+)'
                    all_times = re.findall(time_pattern, all_splits_text)

                    # Parse splits - the pattern is:
                    # r:+0.66 split_time cumulative_time (diff)
                    # OR
                    # r:+0.66 split_time_only (with no parentheses after, meaning no cumulative)
                    #         cumulative_time (diff)

                    # Strategy: Check if each time has parentheses immediately after it
                    # - If YES with format "time (something)" → time is cumulative
                    # - If NO → time is just a split time, no cumulative yet

                    split_num = 1
                    times_idx = 0

                    # Skip reaction time if present
                    if times_idx < len(all_times) and float(all_times[times_idx]) < 1.0:
                        times_idx += 1

                    # Parse times by checking if they have parentheses after them
                    while times_idx < len(all_times) and split_num <= 33:
                        if times_idx >= len(all_times):
                            break

                        current_time = all_times[times_idx]

                        # Determine distance based on split number
                        # For 1650, typically 50 yard splits
                        distance = split_num * 50

                        # Check if current_time has parentheses after it in the original text
                        # Pattern: current_time (something)
                        has_parentheses_after = re.search(
                            re.escape(current_time) + r'\s*\(',
                            all_splits_text
                        )

                        if has_parentheses_after:
                            # This time has parentheses, so it's a cumulative time
                            # We need to find what the split time was
                            # The split time is the value in the parentheses
                            # But we already have it as the "next" time in our list

                            # Actually, let's reconsider:
                            # Format is: cumulative_time (split_diff)
                            # So current_time is the cumulative
                            # We need to calculate or get the split from parentheses

                            # Extract the diff from parentheses
                            paren_match = re.search(
                                re.escape(current_time) + r'\s*\(([\d:.]+)\)',
                                all_splits_text
                            )

                            if paren_match:
                                split_diff = paren_match.group(1)

                                splits_data[f'split_{split_num}_distance'] = distance
                                splits_data[f'split_{split_num}_time'] = split_diff
                                splits_data[f'split_{split_num}_cumulative'] = current_time

                                # Skip the diff time in our times array since we already used it
                                times_idx += 1
                                if times_idx < len(all_times) and all_times[times_idx] == split_diff:
                                    times_idx += 1

                                split_num += 1
                            else:
                                # Couldn't find parentheses pattern, skip this time and move on
                                times_idx += 1
                        else:
                            # This time has NO parentheses after it
                            # So it's just a split time with no cumulative
                            splits_data[f'split_{split_num}_distance'] = distance
                            splits_data[f'split_{split_num}_time'] = current_time
                            splits_data[f'split_{split_num}_cumulative'] = None

                            times_idx += 1
                            split_num += 1

                if swimmer_name:
                    result = {
                        'meet_name': meet_name,
                        'meet_url': meet_url,
                        'event_number': event_number,
                        'event_name': event_name,
                        'is_relay': False,
                        'Rank': rank,
                        'Name': swimmer_name,
                        'Year': year,
                        'School': school,
                        'Finals_Time': finals_time
                    }

                    # Add all split columns (up to 33)
                    for split_idx in range(1, 34):
                        result[f'split_{split_idx}_distance'] = splits_data.get(f'split_{split_idx}_distance', None)
                        result[f'split_{split_idx}_time'] = splits_data.get(f'split_{split_idx}_time', None)
                        result[f'split_{split_idx}_cumulative'] = splits_data.get(f'split_{split_idx}_cumulative', None)

                    results.append(result)
            else:
                i += 1

        print(f"DEBUG: Completed parsing. Found {len(results)} total results")
        return results

    def parse_event_page(self, url, meet_name=None, meet_url=None):
        """
        Parse an event results page and extract all relevant data.

        Args:
            url: URL of the event page to parse
            meet_name: Optional meet name (will be extracted if not provided)
            meet_url: Optional meet URL (will use the full event URL if not given)

        Returns:
            tuple: (pandas.DataFrame, event_type) where event_type is 'relay', 'individual', or 'diving'
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

        # Use the full event URL as meet_url if not provided
        if not meet_url:
            meet_url = url

        # Extract event information
        event_number, event_name, is_relay = self._extract_event_info(page_text)

        if not event_number or not event_name:
            print(f"Could not extract event information from {url}")
            return pd.DataFrame(), None

        print(f"Event {event_number}: {event_name} (Relay: {is_relay})")

        # Check if this is a diving event
        is_diving = 'Diving' in event_name

        # Parse results based on event type
        if is_relay:
            results = self._parse_relay_results(page_text, meet_name, meet_url,
                                                event_number, event_name)
            event_type = 'relay'
        elif is_diving:
            results = self._parse_diving_results(page_text, meet_name, meet_url,
                                                 event_number, event_name)
            event_type = 'diving'
        else:
            results = self._parse_individual_results(page_text, meet_name, meet_url,
                                                     event_number, event_name)
            event_type = 'individual'

        print(f"Extracted {len(results)} results")

        # Convert to DataFrame
        df = pd.DataFrame(results)
        return df, event_type

    def scrape_entire_meet(self, index_url, output_file='meet_results.xlsx'):
        """
        Scrape all events from a meet and save to Excel with separate sheets for relays, individuals, and diving.

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

        # Separate results by type
        relay_results = []
        individual_results = []
        diving_results = []

        # Parse each event
        for i, session in enumerate(sessions):
            print(f"\nProcessing event {i + 1}/{len(sessions)}: {session['event_name']}")

            try:
                df, event_type = self.parse_event_page(session['full_url'],
                                                       meet_name=meet_name,
                                                       meet_url=session['full_url'])

                if not df.empty:
                    if event_type == 'relay':
                        relay_results.append(df)
                    elif event_type == 'individual':
                        individual_results.append(df)
                    elif event_type == 'diving':
                        diving_results.append(df)

                # Be respectful with delays
                time.sleep(self.delay)

            except Exception as e:
                print(f"Error parsing {session['full_url']}: {e}")
                import traceback
                traceback.print_exc()
                continue

        # Save to Excel with multiple sheets
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            if relay_results:
                relay_df = pd.concat(relay_results, ignore_index=True)
                print(f"\nSaving {len(relay_df)} relay results to 'Relay Results' sheet")
                relay_df.to_excel(writer, sheet_name='Relay Results', index=False)

            if individual_results:
                individual_df = pd.concat(individual_results, ignore_index=True)
                print(f"Saving {len(individual_df)} individual results to 'Individual Results' sheet")
                individual_df.to_excel(writer, sheet_name='Individual Results', index=False)

            if diving_results:
                diving_df = pd.concat(diving_results, ignore_index=True)
                print(f"Saving {len(diving_df)} diving results to 'Diving Results' sheet")
                diving_df.to_excel(writer, sheet_name='Diving Results', index=False)

            print(f"\nSuccessfully saved to {output_file}")

        # Return combined results
        all_results = []
        if relay_results:
            all_results.extend(relay_results)
        if individual_results:
            all_results.extend(individual_results)
        if diving_results:
            all_results.extend(diving_results)

        if all_results:
            return pd.concat(all_results, ignore_index=True)
        else:
            print("No results found!")
            return pd.DataFrame()

    def close(self):
        """Close the Selenium driver."""
        if hasattr(self, 'driver'):
            self.driver.quit()


if __name__ == "__main__":
    # Initialize scraper
    scraper = SwimMeetScraper(delay=1.0,  # General delay b/w requests
                              rand_delay_min=8,  # Min random delay b/w split parses
                              rand_delay_max=14,  # Max random delay b/w split parses
                              headless=False  # Set to False to see browser window
                              )

    try:
        # Example: Parse a single individual event
        # event_url = "https://swimmeetresults.tech/NCAA-Division-I-Men-2025/250326F015.htm"
        # df, event_type = scraper.parse_event_page(event_url,
        #                               meet_name="2025 NCAA Division I Men's Swimming & Diving",
        #                               meet_url="https://swimmeetresults.tech/NCAA-Division-I-Men-2025/")

        # print(f"\nEvent Type: {event_type}")
        # print("\nResults preview:")
        # print(df.head().to_string())

        # # Save single event
        # df.to_excel('output_stuff\\single_individual_event.xlsx', sheet_name='Event Results', index=False)
        # print("\nSaved to output_stuff\\single_individual_event.xlsx")

        # Scrape entire meet - will create separate sheets for relay, individual, and diving events
        index_url = "https://swimmeetresults.tech/NCAA-Division-I-Men-2025/index.htm"
        full_results = scraper.scrape_entire_meet(index_url, output_file='ncaa_meet_results.xlsx')

    finally:
        scraper.close()
