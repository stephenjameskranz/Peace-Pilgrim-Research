# Map build plan

Goal: an interactive map of Peace Pilgrim's documented locations, 1952–1981, for moop.earth. Built in small steps, each one usable on its own.

## Approach

- **A plain web page** (HTML + JavaScript) in this repo, under `map/`. No build tools, so it opens straight from GitHub Pages and can be copied onto moop.earth later.
- **Leaflet** for the map: small, simple, well documented, and good at dots, popups and lines.
- **Data comes from the files we already generate.** The page reads `peace-pilgrim-events.geojson`; `scripts/build-all.ps1` keeps it current. Routes drawn later are saved as their own GeoJSON file.
- **Look:** greyscale with one blue accent. Light grey basemap (CARTO Positron), dark grey text and outlines, blue for the strongest evidence and for the selected item. Clean sans-serif type, generous white space.

## Steps

### 1. Basic map
- Greyscale basemap centred on North America.
- One dot per located event, all the same style.
- Click a dot: a popup with the date range, place, what happened, notes, and links to each source clipping (headline, newspaper, link to the scan on peacepilgrim.org).
- Data change needed first: add the source headlines, newspapers and URLs to the GeoJSON, so the popup doesn't have to look them up.
- Publish with GitHub Pages so it can be checked from any computer.

### 2. Dot styling that shows evidence quality
- **Colour = certainty:** blue = Known; mid grey = Known place / approx. date; light grey = Probable.
- **Ring vs filled dot = place precision:** filled for a town or landmark; hollow ring for state or region centres (approximate points).
- **Size = date precision:** small for an exact day, larger for a week, month or year window.
- A small legend explaining all three.
- Several events at one town: group them so they don't stack invisibly (a numbered dot that spreads out on click).

### 3. Time
- A year slider or timeline under the map: show one year, a range, or everything.
- Labels for the seven pilgrimages (1953, 1954–57, and so on) as quick filters.
- Optional play button that steps through time.
- A side list of the events currently shown, in date order; clicking an entry zooms to its dot.

### 4. Straight-line path
- Join the events in date order with straight lines, split into separate pilgrimages.
- Solid line between two Known points; dashed where either end is approximate or a gap is long.
- A toggle to show only Known points in the path.

### 5. Route drawing mode
- An edit mode (for us, not public visitors) to draw a probable route between dots: click waypoints, drag to adjust, label it straight-line or road-following, set its certainty and the sources behind it.
- **Road-following:** ask a free routing service (OSRM or OpenRouteService, walking profile) for the path between waypoints. Today's roads only approximate 1950s highways, so mark these as approximate. Where an article names a highway (US 80, US 77, Route 66, Hwy 101), add waypoints along it to force the route onto that road.
- Save routes to `data/routes.geojson` (download from the page, then commit), so the public map shows them read-only.

### 6. Polish and publish
- Mobile layout, keyboard access, page title and intro text, credits (Friends of Peace Pilgrim archive, OpenStreetMap, CARTO).
- Put it on moop.earth: either embed the GitHub Pages map or copy the `map/` folder into the site.

## To decide later
- How moop.earth is hosted (which affects step 6).
- Whether the route editor stays private (local only) or gets a hidden URL.
