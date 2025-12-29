# SwimcloudScraper
Data scraper for swimcloud

Goal is to pull times and splits from various teams for use in further statistical analysis

So far if you input a team ID (visible publicly on swimcloud), then it will parse a maximum number of meets in the 2024-2025 season (user-defined max) and save all events/swimmers/final times to an excel table, with tabs for each meet.

Current version:
Can do:
  --parse user-defined # of meets
  --identify events within the meet
  --identify teams/individuals within each event
  --selenium integration confirmed to be able to parse split table
  --print to excel output for future datamining

Need to update/fix:
  --selenium parsing works for relays but not individual events (yet)
  --associate swimmer names with their split legs in dictionary for relays
  --different logic relay vs individual event for split handling?
  --ensure compatibility with parsing relay events with > 4 splits (400 yard = 8 splits for 4 swimmers, 800 yard = 16 splits for 4 swimmers)
  be nice to host site to avoid rate limiting...we want to be doing ethical scraping :)
