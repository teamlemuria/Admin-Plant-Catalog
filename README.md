# Admin-Plant-Catalog

This is a lightweight local application for editing and searching the supplied plant-index.xlsx workbook.

## What it does

- Searches all spreadsheet fields with one simple keyword search.
- Displays matching rows.
- Opens any row in a user-friendly edit form.
- Adds new rows through the same form.
- Writes changes directly back to `plant-index.xlsx`.
- Shows results 50 per page, with colour-coded stock status.
- Groups the edit form into Identification, Inventory & Pricing, and Growing Information.
- Suggests existing values (type, list, exposure, etc.) as you type; STOCK NOTE is a dropdown fed by Sheet2.
- Keeps the workbook's STOCK NOTE dropdown intact when saving.
- Click any common name to open that plant's landing page: photo, growing facts, and a table of available sizes.
- Photos come from the `photos` folder; the IMAGE column holds the path inside it (e.g. roses/knockout-red.jpg).
- "Have it" checkbox on every search-result row: tick it and today's date is saved to LAST CHECK right away (no need to open the record). Untick to undo. Ticked rows turn light green.
- Runs only on this computer; it does not require a database or network connection.

## First-time setup

1. Install Python 3.11 or newer if it is not already installed.
2. Open a terminal in this folder.
3. Install the two small dependencies:

   `python -m pip install -r requirements.txt`

4. Start the application:

   `python app.py`

5. Your browser should open automatically at:
   `http://127.0.0.1:5000`

## Important

Keep a backup copy of `plant-index.xlsx` before using the application with the real shared workbook.

The first version intentionally keeps the interface simple and uses the workbook's first worksheet and first row as the column headings.


## Plant landing pages and photos

1. Add a column headed IMAGE to the first row of Sheet1 in plant-index.xlsx (close Excel first).
2. Put pictures in the `photos` folder (sub-folders are fine).
3. Edit a plant and type the picture's path in the Photo section; a preview appears as you type.

One landing page covers every row (size) of the same plant, matched on BOTANICAL NAME + COMMON NAME.
The page uses the first photo found among those rows.

Settings at the top of app.py: PHOTO_DIR, IMAGE_FIELD, STAFF_FIELDS, and SHOW_STAFF_DETAILS
(set it to False before sharing pages with customers: hides internal columns, Edit buttons, and photo warnings).

# Admin-Plant-Catalog

This is a lightweight local application for editing and searching the supplied plant-index.xlsx workbook.

## What it does

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
- Runs only on this computer; it does not require a database or network connection.

## First-time setup

1. Install Python 3.11 or newer if it is not already installed.
2. Open a terminal in this folder.
3. Install the two small dependencies:

   `python -m pip install -r requirements.txt`

4. Start the application:

   `python app.py`

5. Your browser should open automatically at:
   `http://127.0.0.1:5000`

## Important

Keep a backup copy of `plant-index.xlsx` before using the application with the real shared workbook.

The first version intentionally keeps the interface simple and uses the workbook's first worksheet and first row as the column headings.


## Plant landing pages and photos

1. Add a column headed IMAGE to the first row of Sheet1 in plant-index.xlsx (close Excel first).
2. Put pictures in the `photos` folder (sub-folders are fine).
3. Edit a plant and enter one or more picture paths in the Photo section, separated by semicolons; previews appear as you type. Existing single-photo entries continue to work.

One landing page covers every row (size) of the same plant, matched on BOTANICAL NAME + COMMON NAME.
The page uses the first photo found among those rows.

Settings at the top of app.py: PHOTO_DIR, IMAGE_FIELD, STAFF_FIELDS, and SHOW_STAFF_DETAILS
(set it to False before sharing pages with customers: hides internal columns, Edit buttons, and photo warnings).
