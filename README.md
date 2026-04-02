# Excel Report Generator

A Python-based automated report generation tool that transforms raw data into professional, formatted Excel reports with charts, pivot tables, and custom styling.

## Features

- **Automated Formatting**: Professional styling with headers, borders, and color schemes
- **Dynamic Charts**: Auto-generated charts and visualizations
- **Multiple Data Sources**: CSV, JSON, SQL databases, APIs
- **Template System**: Reusable report templates
- **Data Validation**: Built-in quality checks before report generation
- **Scheduling Ready**: Can be automated with cron/Task Scheduler

## Use Cases

- Daily/weekly business reports
- Financial summaries and dashboards
- Sales performance reports
- Data exports from internal systems
- KPI tracking and monitoring

## Tech Stack

- **Python 3.8+**
- **openpyxl** - Excel file manipulation
- **pandas** - Data processing
- **xlsxwriter** - Advanced Excel features
- **matplotlib/seaborn** - Chart generation

## Project Structure

```
excel-report-generator/
├── report_generator.py   # Main report generation engine
├── data_loader.py        # Load data from various sources
├── formatters.py         # Excel styling and formatting
├── templates/            # Report templates
│   ├── sales_report.py
│   └── financial_summary.py
├── sample_data/          # Sample datasets
│   └── sample_sales.csv
├── config.py             # Configuration
├── requirements.txt
└── README.md
```

## Setup

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd excel-report-generator
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run sample report**
   ```bash
   python report_generator.py --template sales --input sample_data/sample_sales.csv
   ```

## Usage

### Basic Usage
```bash
python report_generator.py --input data.csv --output report.xlsx
```

### With Template
```bash
python report_generator.py --template financial --input data.csv --output monthly_report.xlsx
```

### From Database
```bash
python report_generator.py --source database --query "SELECT * FROM sales" --output report.xlsx
```

## Report Templates

### Sales Report
- Summary statistics
- Sales by region/product
- Trend charts
- Top performers

### Financial Summary
- Revenue breakdown
- Expense analysis
- Profit/loss statement
- Budget vs actual

### Custom Template
Create your own template in `templates/` directory.

## Features in Detail

### Auto-Formatting
- Professional headers with company colors
- Alternating row colors for readability
- Auto-fit columns
- Number formatting (currency, percentages, dates)
- Conditional formatting for KPIs

### Charts & Visualizations
- Bar charts
- Line charts
- Pie charts
- Combo charts
- Sparklines

### Data Validation
- Missing value checks
- Data type validation
- Range validation
- Duplicate detection

## Configuration

Edit `config.py`:
```python
COMPANY_NAME = "Your Company"
PRIMARY_COLOR = "#4472C4"
SECONDARY_COLOR = "#ED7D31"
CHART_STYLE = 10
```

## Scheduling

### Windows Task Scheduler
```bash
# Create a batch file
python C:\path\to\report_generator.py --template sales --output daily_report.xlsx
```

### Linux Cron
```bash
0 8 * * 1 /usr/bin/python3 /path/to/report_generator.py --template weekly
```

## Examples

See `examples/` directory for:
- Sales report example
- Financial dashboard example
- Custom formatting examples

## Future Enhancements

- [ ] PDF export option
- [ ] Email integration (auto-send reports)
- [ ] Multiple sheet support
- [ ] Pivot table automation
- [ ] Web dashboard integration
- [ ] Real-time data refresh

## Author

Built to demonstrate automation skills and data reporting capabilities for business analyst roles.

## License

MIT License
