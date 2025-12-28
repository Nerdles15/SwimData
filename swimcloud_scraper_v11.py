# Based on v12, dynamic output filename from team name

from selenium import webdriver
import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
from urllib.parse import urljoin
import re

class SwimCloudScraper:
    def __init__(self, delay=1.0):
        """
        Initialize the scraper with a delay between requests.
        
        Args:
            delay: Seconds to wait between requests (default 1.0)
        """
        self.base_url = "https://www.swimcloud.com"
        self.delay = delay
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        self.team_name = None
    
    def _delay_request(self):
        """Add delay between requests to be respectful to the server."""
        time.sleep(self.delay)
    
    def get_team_name(self, team_id):
        """
        Get the team name from the team page.
        
        Args:
            team_id: The team ID (e.g., 185)
        
        Returns:
            String with team name
        """
        url = f"{self.base_url}/team/{team_id}/results/?page=1&name=&meettype=&season=28"
        print(f"Fetching team name from: {url}")
        
        try:
            response = self.session.get(url)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Look for h1 with class c-toolbar__title
            team_name_tag = soup.find('h1', class_='c-toolbar__title')
            if team_name_tag:
                team_name = team_name_tag.get_text(strip=True)
                print(f"Found team name: {team_name}")
                return team_name
            else:
                print("WARNING: Could not find team name")
                return f"Team_{team_id}"
        
        except Exception as e:
            print(f"Error fetching team name: {e}")
            return f"Team_{team_id}"
    
    def get_team_meets(self, team_id, max_meets=None):
        """
        Get all meet URLs for a given team.
        
        Args:
            team_id: The team ID (e.g., 185)
            max_meets: Maximum number of meets to retrieve (None for all)
        
        Returns:
            List of meet URLs
        """
        url = f"{self.base_url}/team/{team_id}/results/?page=1&name=&meettype=&season=28"
        print(f"Fetching team results from: {url}")
        
        try:
            response = self.session.get(url)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Find all meet links
            meet_links = []
            all_links = soup.find_all('a', href=True)
            
            for link in all_links:
                href = link['href']
                # Look for any link containing /results/ followed by digits
                if '/results/' in href and re.search(r'/results/(\d+)', href):
                    # Extract just the meet result URL
                    match = re.search(r'(/results/\d+)/?', href)
                    if match:
                        clean_path = match.group(1) + '/'
                        meet_url = urljoin(self.base_url, clean_path)
                        if meet_url not in meet_links:
                            meet_links.append(meet_url)
            
            if not meet_links:
                print("WARNING: No meet links found!")
                print("Saving HTML for debugging...")
                with open(f'team_{team_id}_debug.html', 'w', encoding='utf-8') as f:
                    f.write(soup.prettify())
                print(f"Saved page HTML to team_{team_id}_debug.html")
            
            if max_meets:
                meet_links = meet_links[:max_meets]
            
            print(f"Found {len(meet_links)} meets")
            return meet_links
        
        except Exception as e:
            print(f"Error fetching team meets: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def get_meet_events(self, meet_url):
        """
        Get meet name and all event URLs for a given meet.
        
        Args:
            meet_url: URL of the meet
        
        Returns:
            Tuple of (meet_name, list of tuples (event_url, event_number, event_name))
        """
        print(f"\nFetching events from meet: {meet_url}")
        self._delay_request()
        
        try:
            response = self.session.get(meet_url)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Extract meet name
            meet_name = "Unknown Meet"
            meet_name_tag = soup.find('h1', id='meet-name')
            if not meet_name_tag:
                meet_name_tag = soup.find('h1', class_='c-toolbar__title')
            if meet_name_tag:
                meet_name = meet_name_tag.get_text(strip=True)
            
            # Extract meet ID from URL
            match = re.search(r'/results/(\d+)', meet_url)
            if not match:
                print("Could not extract meet ID from URL")
                return meet_name, []
            
            meet_id = match.group(1)
            
            # Find all event links with their names
            event_links = []
            all_links = soup.find_all('a', href=True)
            
            for link in all_links:
                href = link['href']
                # Match patterns like /results/307921/event/11/
                match = re.match(rf'^/results/{meet_id}/event/(\d+)/?$', href)
                if match:
                    event_number = match.group(1)
                    event_url = urljoin(self.base_url, f'/results/{meet_id}/event/{event_number}/')
                    
                    # Extract event name from the div.c-events__link-body
                    event_name = "Unknown Event"
                    event_body = link.find('div', class_='c-events__link-body')
                    if event_body:
                        # Try title attribute first
                        if event_body.get('title'):
                            event_name = event_body.get('title')
                        else:
                            # Otherwise get text content
                            event_name = event_body.get_text(strip=True)
                    
                    if (event_url, event_number, event_name) not in event_links:
                        event_links.append((event_url, event_number, event_name))
            
            if not event_links:
                print("WARNING: No event links found in meet!")
                print("Saving HTML for debugging...")
                with open(f'meet_{meet_id}_debug.html', 'w', encoding='utf-8') as f:
                    f.write(soup.prettify())
                print(f"Saved meet HTML to meet_{meet_id}_debug.html")
            
            print(f"Meet: {meet_name}")
            print(f"Found {len(event_links)} events")
            return meet_name, event_links
        
        except Exception as e:
            print(f"Error fetching meet events: {e}")
            import traceback
            traceback.print_exc()
            return "Unknown Meet", []

    def scrape_split_times(self, time_url):
        driver = webdriver.Chrome()
        driver.get(time_url)

        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")

        table = soup.select_one("table.c-table-clean")

        trs = table.find_all("tr")

        for tr in trs:
            # print(tr.get_text(" | ", strip=True)) ## Just printing the data to see if I was grabbing the corect data.

            data = tr.get_text()

            pattern = r'(\d+?)(\d{2}\.\d{2})(\d{2}\.\d{2})(\d+:\d{2}\.\d{2}|\d+\.\d{2})$'

            match = re.match(pattern, data)

            if match:
                values = match.groups()
                print(values)
            else:
                print("no match! line doesn't contain numbers, need to ignore that.")


    
    def get_event_results(self, event_url, event_name):
        """
        Get all results (names and times) for a specific event.
        
        Args:
            event_url: URL of the event
            event_name: Name of the event (already extracted from meet page)
        
        Returns:
            Dictionary with event_name, is_relay flag, and list of results
        """
        print(f"  Fetching results for: {event_name}")
        self._delay_request()
        
        try:
            response = self.session.get(event_url)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Check if this is a relay event
            is_relay = 'relay' in event_name.lower()
            
            # Find all result entries
            results = []
            
            # Strategy: Find all divs with id starting with "time" which contain the time links
            # Example: <div id="time148087775"><a href="/times/148087775/">1:35.48</a></div>
            time_divs = soup.find_all('div', id=re.compile(r'^time\d+'))
            
            for time_div in time_divs:
                # Extract time value from the link inside the div
                time_link = time_div.find('a', href=re.compile(r'^/times/\d+'))
                if not time_link:
                    continue
                
                time_value = time_link.get_text(strip=True)
                time_url = urljoin(self.base_url, time_link['href'])
                # Now find the corresponding athlete/team name
                # Look for the nearest td with class="u-nowrap u-text-semi" that has a swimmer link
                # We need to traverse up and find the row, then look for the name
                self.scrape_split_times(time_url)
                # Find the parent table row
                row = time_div.find_parent('tr')
                name = "Unknown"
                
                if row:
                    if is_relay:
                        # For relays, look for team link
                        team_link = row.find('a', href=re.compile(r'/team/\d+'))
                        if team_link:
                            name = team_link.get_text(strip=True)
                    else:
                        # For individuals, look for swimmer link
                        swimmer_link = row.find('a', href=re.compile(r'/swimmer/\d+'))
                        if swimmer_link:
                            name = swimmer_link.get_text(strip=True)
                            name = re.sub(r'\s+', ' ', name)
                
                results.append({
                    'name': name,
                    'time': time_value,
                    'time_url': time_url
                })
            
            print(f"    Found {len(results)} results | Relay: {is_relay}")
            
            return {
                'event_name': event_name,
                'is_relay': is_relay,
                'results': results
            }
        
        except Exception as e:
            print(f"  Error fetching event results: {e}")
            import traceback
            traceback.print_exc()
            return {'event_name': event_name, 'is_relay': False, 'results': []}
    
    def scrape_team_results(self, team_id, max_meets=None, output_file=None):
        """
        Scrape all results for a team and save to CSV.
        
        Args:
            team_id: The team ID
            max_meets: Maximum number of meets to scrape (None for all)
            output_file: Output Excel filename (None to auto-generate from team name)
        
        Returns:
            DataFrame with all results
        """
        print(f"\n{'='*70}")
        print(f"Starting scrape for Team ID: {team_id}")
        print(f"Max meets: {max_meets if max_meets else 'All'}")
        print(f"{'='*70}\n")
        
        # Get team name
        self.team_name = self.get_team_name(team_id)
        
        # Generate output filename from team name if not provided
        if output_file is None:
            # Clean team name for filename (remove special characters)
            clean_name = re.sub(r'[^\w\s-]', '', self.team_name)
            clean_name = re.sub(r'[-\s]+', '_', clean_name)
            output_file = f'{clean_name}.xlsx'
        
        print(f"Output file: {output_file}\n")
        
        df = pd.DataFrame()
        
        # Get all meets for the team
        meet_urls = self.get_team_meets(team_id, max_meets)
        
        if not meet_urls:
            print("\n❌ No meets found. Please check the team ID or page structure.")
            return pd.DataFrame()
        
        # Create a new workbook
        with pd.ExcelWriter(output_file) as writer:
            df.to_excel(writer, index=False)
        
        for meet_idx, meet_url in enumerate(meet_urls, 1):
            all_results = []
            print(f"\n{'─'*70}")
            print(f"Processing Meet {meet_idx}/{len(meet_urls)}")
            print(f"{'─'*70}")
            
            # Get meet name and all events in the meet
            meet_name, event_links = self.get_meet_events(meet_url)
            
            if not event_links:
                print(f"  ⚠️  No events found in this meet, skipping...")
                continue
            
            for event_url, event_number, event_name in event_links:
                # Get all results for this event directly from the event page
                event_data = self.get_event_results(event_url, event_name)
                is_relay = event_data['is_relay']
                results = event_data['results']
                
                if not results:
                    print(f"    ⚠️  No results found for event {event_number}")
                    continue
                
                # Add each result to our data
                for result in results:
                    all_results.append({
                        'meet_name': meet_name,
                        'meet_url': meet_url,
                        'event_number': event_number,
                        'event_name': event_name,
                        'is_relay': is_relay,
                        'name': result['name'],
                        'time': result['time'],
                        'time_url': result['time_url']
                    })
                print(f"      Added {len(results)} results")
            print(f"{'─'*70}\nCompleted Meet: {meet_name}\n{'─'*70}")

            # Create DataFrame for this meet and append to Excel
            df_meet = pd.DataFrame(all_results)
            if not df_meet.empty:
                # Truncate string for sheet name compatibility
                trunc_meet_name = df_meet['meet_name'].apply(lambda x: x[:31] if len(x) > 31 else x)
                # Append meet data to existing Excel file as new sheets
                with pd.ExcelWriter(output_file, mode='a') as writer:
                    df_meet.to_excel(writer, sheet_name=trunc_meet_name.iloc[0], index=False)
                print(f"\n{'='*70}")
                print(f"✅ Saved {len(df_meet)} results for meet '{trunc_meet_name.iloc[0]}' to {output_file}")
                print(f"{'='*70}\n")
                df = pd.concat([df, df_meet], ignore_index=True)
            else:
                print(f"\n❌ No results found for meet '{meet_name}'.\n")

        print(f"\n{'='*70}")
        print(f"✅ Scraping complete!")
        print(f"   Team: {self.team_name}")
        print(f"   Saved {len(df)} results to {output_file}")
        print(f"   Meets processed: {df['meet_name'].nunique()}")
        print(f"   Unique events: {df['event_name'].nunique()}")
        print(f"   Relay results: {df['is_relay'].sum()}")
        print(f"   Individual results: {(~df['is_relay']).sum()}")
        print(f"{'='*70}\n")

        # Crate our array that will be a data frame to go to excel
        all_results = []

        # Loop through all the links to
        # for result in results:
        #     print()
        #     self.scrape_split_times(result['time_url'])


        return df


# Example usage
if __name__ == "__main__":
    # Initialize scraper with 1 second delay between requests
    scraper = SwimCloudScraper(delay=1.0)
    
    # Scrape results for team 5245, limiting to 2 meets for testing
    # Remove or increase max_meets for production use
    # output_file will be auto-generated from team name
    results_df = scraper.scrape_team_results(
        team_id=5245,
        max_meets=2,  # Set to None to scrape all meets
        output_file=None  # Set to None to auto-generate filename from team name
    )
    
    # Display sample of results
    if not results_df.empty:
        print(f"\n📈 Summary:")
        print(f"   Total results: {len(results_df)}")
        print(f"   Meets: {results_df['meet_name'].nunique()}")
        print(f"   Events: {results_df['event_name'].nunique()}")