# Lemuria Nursery Plant Catalog (combined app)

This single local application combines the Admin Plant Catalog and the static public-site builder. It edits the Excel workbook and builds the public website from that same workbook and the same `photos` folder. No copying files between applications is needed.

It is a lightweight local application for editing and searching the supplied `plant-index.xlsx` workbook. It runs only on this computer; it does not require a database or network connection.

## Start the application

1. Keep `plant-index.xlsx` closed in Excel while using the app.
2. Double-click `Start Plant Index.bat`. On first use, it installs the required Python packages.
3. Your browser opens the local admin interface at `http://127.0.0.1:5000`.

### Manual setup (alternative)

1. Install Python 3.11 or newer if it is not already installed.
2. Open a terminal in this folder.
3. Install the two small dependencies:

   `python -m pip install -r requirements.txt`

4. Start the application:

   `python app.py`

5. Your browser should open automatically at:
   `http://127.0.0.1:5000`

## What the admin catalog does

- Searches all spreadsheet fields with one simple keyword search.
- Displays matching rows.
- Opens any row in a user-friendly edit form.
- Adds new rows through the same form.
- Writes changes directly back to `plant-index.xlsx`.
- Shows results 50 per page, with colour-coded stock status.
- Change STOCK NOTE directly from the inventory table; the chosen status saves to Excel immediately.
- STOCK NOTE colors: In stock (green), Low stock (blue), Check stock (yellow), Not available (red).
- Groups the edit form into Identification, Inventory & Pricing, and Growing Information.
- Suggests existing values (type, list, exposure, etc.) as you type; STOCK NOTE is a dropdown fed by Sheet2.
- Keeps the workbook's STOCK NOTE dropdown intact when saving.
- Click any common name to open that plant's landing page: photo, growing facts, and a table of available sizes.
- Photos come from the `photos` folder; the IMAGE column holds one or more paths inside it (e.g. `roses/knockout-red.jpg; roses/knockout-flower.jpg`).
- "Have it" checkbox on every search-result row: tick it and today's date is saved to LAST CHECK right away (no need to open the record). Untick to undo. Ticked rows turn light green.

## Maintain the catalog

Use the app to search records, update inventory, edit records, add plants, and manage photos. Changes are saved to `plant-index.xlsx`; photos are stored in `photos/`.

The application uses the workbook's first worksheet and first row as the column headings. The interface is intentionally kept simple.

## Plant landing pages and photos

1. Add a column headed IMAGE to the first row of Sheet1 in `plant-index.xlsx` (close Excel first).
2. Put pictures in the `photos` folder (sub-folders are fine).
3. Edit a plant and enter one or more picture paths in the Photo section, separated by semicolons; previews appear as you type. Existing single-photo entries continue to work.

One landing page covers every row (size) of the same plant, matched on BOTANICAL NAME + COMMON NAME. The page uses the first photo found among those rows.

Settings at the top of `app.py`: `PHOTO_DIR`, `IMAGE_FIELD`, `STAFF_FIELDS`, and `SHOW_STAFF_DETAILS` (set it to `False` before sharing pages with customers: hides internal columns, Edit buttons, and photo warnings).

## Update the public website

1. Save any changes to the plant catalog.
2. Click **Update site** on the Plant Inventory page.
3. Wait for the completion message. The static website is generated in `public-site/`.
4. Open `public-site/index.html` in a browser to preview the generated site. Upload the entire `public-site/` folder to your web host when ready.

The site builder reads `plant-index.xlsx` and `photos/` directly. It does not change the workbook. The generated public site hides staff-only fields according to `site.json` and the existing public-site rules. Clicking **Update site** only creates local files; it does not publish them to the internet.

## Website text and display settings

Edit `site.json` in a text editor to change the public site's name, introduction, about/visit text, contact details, hours, availability labels, categories, and featured-plant count. The site templates and styling are in `site_templates/` and `site_static/`.

## Backup

Back up `plant-index.xlsx` and the `photos/` folder regularly. Keep a backup copy of `plant-index.xlsx` before using the application with the real shared workbook. The generated `public-site/` folder can be rebuilt from those source files and `site.json`.
