# Weekly Client Dispatch Board

A responsive Dash web application for managing and visualizing weekly client work orders and dispatch scheduling. The application reads data from Excel spreadsheets and presents it in an interactive, filterable table with intelligent flagging for scheduling conflicts and priority indicators.

## Features

### 📊 Interactive Dashboard
- **Dynamic Table View**: Displays work orders organized by company, region, and day of week
- **Real-time Filtering**: Filter by week, work order state, company, region, and custom flags
- **Sticky Columns**: Company and region information remain visible while scrolling
- **Synchronized Scrolling**: Coordinated horizontal scrolling with top indicator bar
- **Drag-to-Scroll**: Intuitive mouse drag for horizontal navigation

### 🚩 Smart Flag Detection
Automatic detection and visual highlighting of:
- **Consecutive Days**: Same company/location work on consecutive days
- **Same Site**: Multiple work orders at same location across multiple days
- **Same Case**: Single case number appearing across multiple days
- **Multi-Group**: Multiple assignment groups for same company
- **Escalated**: Cases marked with escalation status
- **Warning**: Cases with warning indicators
- **RCE**: Remote support requirements detected in notes

### ⚙️ Flexible Filtering
- Week selection with automatic current week detection
- Work order state filtering (e.g., "Assigned")
- Multi-select company and region filters
- Custom flag-based filtering
- Weekend visibility toggle
- "Show only flagged rows" option

### Theme Palette (Default + Canyon + Forest)
- Pill button selector in the header (`Default`, `Canyon`, `Forest`)
- Instant recoloring across header, controls, panels, table, and chips
- Palette selection is retained in app state for the active session
- Default palette restored, with `Canyon` and a higher-contrast `Forest` variant

### 📤 Data Management
- **Upload New Data**: Drag-and-drop Excel file uploads
- **Automatic Validation**: File format and size checks
- **Smart Excel Parsing**: Automatic header detection for variable spreadsheet formats
- **Data Merging**: Combines work order and case data from multiple sheets
- **Real-time Timestamps**: Tracks source file and board refresh times

---

## Installation

### Prerequisites
- Python 3.9+
- pip or conda

### Setup

1. **Clone or Download the Repository**
   ```bash
   cd "Weekly Client dispatch Board"
   ```

2. **Create Virtual Environment** (Recommended)
   ```bash
   python -m venv venv
   # On Windows
   venv\Scripts\activate
   # On macOS/Linux
   source venv/bin/activate
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

### Dependencies
- **pandas** (≥1.5): Data processing and manipulation
- **numpy** (≥1.23): Numerical operations
- **plotly** (≥5.10): Visualization base library
- **dash** (≥2.9): Web framework
- **dash-ag-grid** (≥0.6): Advanced grid component
- **openpyxl** (≥3.0): Excel file reading/writing
- **gunicorn** (≥20.1): Production WSGI server

---

## Usage

### Local Development

#### Run with Default Settings
```bash
python weekly_client_dispatch_board.py
```
- Opens at `http://127.0.0.1:8052`
- Looks for data in `./data` directory

#### Run with Custom Port
```bash
python weekly_client_dispatch_board.py --port 8080
```

#### Run with Specific Excel File
```bash
python weekly_client_dispatch_board.py --case-file path/to/Cases_Final_Dashboard_CURRENT.xlsx
```

#### Debug Mode
```bash
python weekly_client_dispatch_board.py --debug
```

### Production Deployment

#### Using Gunicorn
```bash
gunicorn --worker-class gthread --workers 1 --threads 2 \
  --timeout 120 --bind 0.0.0.0:8000 \
  weekly_client_dispatch_board:server
```

#### Railway Deployment
The included `Procfile` configures the app for Railway:
```
web: gunicorn --worker-class gthread --workers 1 --threads 2 --timeout 120 --max-requests 500 --max-requests-jitter 50 --bind 0.0.0.0:${PORT} weekly_client_dispatch_board:server
```

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `DATA_DIR` | `./data` (or `/data` on Railway) | Directory for Excel file storage |
| `CURRENT_FILE` | `{DATA_DIR}/Cases_Final_Dashboard_CURRENT.xlsx` | Path to current Excel file |
| `APP_TIMEZONE` | `America/New_York` | Timezone for timestamps |
| `CASE_FILE` | None | Path to initial case file at startup |
| `PORT` | `8052` | Local port (set automatically on Railway) |

**Example with Environment Variables:**
```bash
# macOS/Linux
export DATA_DIR=/data/workorders
export APP_TIMEZONE="America/Chicago"
python weekly_client_dispatch_board.py

# Windows PowerShell
$env:DATA_DIR="C:\data\workorders"
$env:APP_TIMEZONE="America/Chicago"
python weekly_client_dispatch_board.py
```

---

## Excel File Format

### Required Spreadsheets

#### Sheet 1: Dispatch Data (auto-detected)
The first sheet should contain work order information with headers like:

| Required Headers | Alternative Names |
|------------------|------------------|
| Work Order | Work order, Number2, WO |
| Case Number | Case number, Number, CS |
| Company | (Standard) |
| Location | (Standard) |
| Work Order State | WO State |
| Assignment Group | Work order Assignment Group, Work Order Assignment Group |
| Scheduled Start | Scheduled start |
| Short Description | Short description, Description |
| State / Province | State |
| City | City / Town, Town |

#### Sheet 2: Detail Sheet (`Detail`) - *Primary source for RCE flags*
Optional sheet for case-level details and RCE flag detection:

| Column Name | Purpose |
|-------------|---------|
| Case Number | Links to dispatch data for case matching |
| Next Steps | Scanned for "RCE" keyword to detect remote support requirements |

**Note**: The Detail sheet is the primary source for RCE flag detection. If present, it will be used. Otherwise, the Per_Case_Dashboard sheet is used as a fallback.

#### Sheet 3: Case Dashboard (`Per_Case_Dashboard`) - *Fallback for additional flags*
Optional sheet for case-level flags (used when Detail sheet is not available):

| Required Headers | Purpose |
|------------------|---------|
| Case Number | Links to dispatch data |
| State / Province | Case location |
| Escalation Status | Detects "escalated" and "warning" keywords |
| Visit Count | Number of visits for case |
| Visit List | (Optional) |
| Summary of case and visits | Scanned for "RCE" or "remote support" keywords |
| Work notes | Scanned for "RCE" or "remote support" keywords |
| Next Steps | (Optional) |

### Data Processing
- Headers are automatically detected (flexible formatting)
- Whitespace and Unicode characters are normalized
- Empty cells are handled gracefully
- Multiple header row formats supported
- Data extracted from rows below headers even if non-standard formatting

**Maximum File Size**: 25 MB

---

## Architecture

### Data Flow
```
Excel File → Normalize Headers → Extract Data → Merge Case Flags → Cache → Render UI
                                                                       ↓
                                                              User Filters → Update Table
```

### Key Components

#### Data Processing (`normalize_dispatch_df`, `normalize_case_flags`)
- Reads Excel with automatic header detection
- Cleans and standardizes column names
- Extracts work orders, case numbers, scheduling info
- Detects escalation and warning flags

#### Caching System
- Thread-safe `RLock`-protected cache
- Prevents repeated Excel parsing
- Updated on file refresh or upload

#### UI Rendering (`build_dispatch_rows`, `render_dispatch_board`)
- Organizes data by company
- Applies filters and grouping
- Generates interactive table rows
- Applies CSS-based visual flags

#### Callback System
- **Refresh Data**: Updates cache from latest file
- **Upload Handling**: Validates and saves new Excel files
- **Board Rendering**: Filters and displays data based on selections
- **Filter Synchronization**: Updates dropdown options based on available data

---

## Customization

### Change Default Timezone
Edit the initialization or set environment variable:
```python
APP_TIMEZONE = ZoneInfo(os.environ.get("APP_TIMEZONE", "America/Los_Angeles"))
```

### Adjust Description Length
```python
SHORT_DESCRIPTION_MAX_LEN = 120  # Characters (default: 160)
```

### Add Custom Flags
Modify the `build_dispatch_rows` function to detect additional patterns:
```python
# Example: detect specific assignment groups
if any("URGENT" in str(grp.get("Assignment Group", "")) for _, grp in ...):
    flags.append("Urgent")
```

### Style Customization
Modify the inline CSS in the `app.index_string` property to change:
- Colors and themes
- Table layout
- Responsive behavior
- Typography

### Built-in Palettes
- Default: original blue board styling
- Canyon: warm tan and neutral tones
- Forest: higher-contrast green earth-tone palette

Switch between palettes using the header pill buttons while the app is running.

---

## Troubleshooting

### "Upload Cases_Final_Dashboard_CURRENT.xlsx to load the board"
- The expected Excel file is not found
- **Solution**: Upload the correct Excel file using the "Upload Dataset" button

### Data Not Updating After Upload
- File may be locked or in use
- **Solution**: Ensure no other applications have the file open; wait 30 seconds and refresh

### Incorrect Data Detected
- Headers not automatically detected
- **Solution**: Rename columns to match expected headers exactly or check sheet name matches `Per_Case_Dashboard` for case data

### Timezone Issues
- Timestamps showing incorrect times
- **Solution**: Set `APP_TIMEZONE` environment variable to your timezone (e.g., `America/Chicago`)

### Table Scrolling Issues
- Sticky columns not aligning
- **Solution**: Refresh browser page; check browser console for JavaScript errors
- Recent updates improved sticky-column left-offset anchoring to reduce overlap/cropping while horizontal scrolling

### Header Says "No workbook loaded" After Browser Refresh
- Board data appears but source status message looks stale
- **Solution**: Updated app layout/render flow now rehydrates header source status from cached active workbook on reload

### Memory/Performance Issues
- Large Excel files (>10,000 rows) causing slowdown
- **Solution**: Increase server resources or filter data before upload

---

## Browser Compatibility

- **Chrome/Edge**: ✅ Full support
- **Firefox**: ✅ Full support
- **Safari**: ✅ Full support (sticky positioning may vary)
- **IE 11**: ❌ Not supported (uses modern CSS)

**Recommended**: Latest versions of Chrome or Edge for optimal performance

---

## Performance Tips

1. **Optimize Excel Files**
   - Remove empty rows/columns before upload
   - Use consistent date formatting

2. **Network**
   - Deploy closer to users if possible
   - Consider caching on frontend for repeat visits

3. **Database Alternative**
   - For >50MB files, consider storing in database instead of Excel

4. **Concurrent Users**
   - Default: 1 worker, 2 threads (Railway free tier)
   - Increase workers/threads for more users in production

---

## Security Considerations

✅ **Implemented**
- File upload validation (extension, size limits)
- XSS protection (Dash auto-escapes)
- Path traversal prevention

⚠️ **Recommendations**
- Add authentication for production deployments
- Use HTTPS in production
- Validate environment variables
- Set appropriate CORS policies if accessed cross-domain

---

## Development

### Project Structure
```
Weekly Client dispatch Board/
├── weekly_client_dispatch_board.py  # Main application
├── requirements.txt                  # Python dependencies
├── Procfile                          # Railway deployment config
├── README.md                         # This file
├── CODE_REVIEW.md                    # Code review and recommendations
└── data/                             # Excel files storage
    └── Cases_Final_Dashboard_CURRENT.xlsx
```

### Adding Features

#### Add a New Filter
1. Add dropdown to toolbar in `build_dispatch_layout`
2. Extract unique values in `dispatch_filter_options`
3. Add filter logic to `build_dispatch_rows`
4. Update callback inputs

#### Add a New Flag
1. Detect condition in `build_dispatch_rows` within company grouping
2. Append to `flags` list
3. Add to `dispatch_flag_options` in `build_dispatch_layout`
4. Add CSS styling for visual indication

---

## Maintenance

### Regular Tasks
- Monitor error logs on production server
- Archive old Excel files to prevent storage bloat
- Update Python packages quarterly: `pip install -U -r requirements.txt`

### Updating Dependencies
```bash
pip list --outdated
pip install --upgrade <package-name>
```

---

## License
*[Add your license here]*

## Contact
For questions or issues, please contact [your contact information].

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | May 2026 | Initial release |

---

## FAQ

**Q: Can I use a CSV file instead of Excel?**
A: Currently only `.xlsx` files are supported. Convert CSV to Excel first.

**Q: How often should I refresh the data?**
A: Click "Refresh Data" when you've updated the source Excel file, or set up periodic automation.

**Q: Can multiple users upload different files?**
A: Only one "current" file is stored. Latest upload replaces previous.

**Q: What happens to old uploaded files?**
A: They're overwritten. Maintain backups separately if needed.

**Q: Can I export the filtered view?**
A: Not built-in, but you can screenshare or export the source Excel with specific worksheets.

---

**Last Updated**: May 7, 2026

