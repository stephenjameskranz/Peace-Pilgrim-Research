# Peace Pilgrim itinerary

A documented record of where and when Peace Pilgrim (Mildred Norman, 1908–1981) was during her pilgrimage, 1953–1981, built to drive a map on moop.earth.

The primary source is the newspaper clipping archive kept by Friends of Peace Pilgrim:
https://www.peacepilgrim.org/articles-about-peace-pilgrim

## Outputs

| File | What it is |
|---|---|
| `peace-pilgrim-itinerary.xlsx` | Workbook with three tabs: About, Location events, Sources. This is what gets imported into the Google Sheet (File → Import → Replace spreadsheet). |
| `peace-pilgrim-location-events.csv` | One row per claim about where she was and when, with dates, certainty, coordinates and source ids. This is the map dataset. |
| `peace-pilgrim-sources.csv` | One row per clipping on the archive page: date, place, headline, newspaper, link, and notes read from the clipping. |
| `peace-pilgrim-events.geojson` | The located events as GeoJSON points, ready for a web map. |

All four are generated. Edit the files in `data/`, then rebuild.

## Data files (edit these)

| File | What it holds |
|---|---|
| `data/links.csv` | Every clipping link scraped from the archive page (title and URL). |
| `data/index_parsed.csv` | The archive index parsed into date / place / headline, with corrections. Generated from `links.csv` by `scripts/build.ps1`; corrections live in that script. |
| `data/reads.csv` | What was read from each clipping: newspaper, miles, route walked and planned, people, notable details, how the date was worked out. Keyed by source id. |
| `data/events_manual.csv` | Hand-made location events drawn from the clippings, often down to a single night. |
| `data/extra_sources.csv` | Clippings linked from the archive page but not listed as their own entries. |
| `data/geocode_cache.csv` | Coordinates for each place, from OpenStreetMap Nominatim. |

## Rebuilding

Windows PowerShell, from the project folder:

```powershell
scripts\build-all.ps1
```

Add `-Geocode` to look up coordinates for any new places first (needs internet; one request per second).

To get the clipping scans on a new computer (they are not stored in the repo):

```powershell
scripts\download-clippings.ps1
```

## How to read the data

- **certainty**: `Known` = a source places her there on that date (or within a day or two). `Known place / approx. date` = the place is documented but the date is only a month, year or window. `Probable` = a place she said she passed through or planned to visit, with the date estimated from route order.
- **date_basis** says where each date comes from. Newspaper publication dates usually trail her visit by a day or two.
- **geo_level** says what a point represents: `city`, `city (first of several)` for rows naming a route or several towns, `state` / `region` for approximate points (draw them as such), or `landmark`.
- Errors in the archive index (wrong states, wrong years, broken or swapped links) are flagged in the `notes` column, never silently changed. `index_title` keeps the original wording.

## Status

- 273 of 281 clippings read. Not read: four index entries with no working link, two photo-only pages from 1986, and two files already read under another entry.
- 402 of 403 events have coordinates. The Appalachian Trail hike (1952) is a route, not a point.

## Copyright

The clippings are copyrighted newspaper articles. The scans stay out of this repository, and the data here records facts and short descriptions, not article text.
