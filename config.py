"""
Configuration for Excel Report Generator
"""
from datetime import datetime

class Config:
    """Report configuration settings"""

    # Company branding
    COMPANY_NAME = "Your Company"
    REPORT_TITLE_PREFIX = "Automated Report"

    # Color scheme (Excel color codes)
    PRIMARY_COLOR = "4472C4"      # Blue
    SECONDARY_COLOR = "ED7D31"    # Orange
    SUCCESS_COLOR = "70AD47"      # Green
    WARNING_COLOR = "FFC000"      # Yellow
    DANGER_COLOR = "FF0000"       # Red
    HEADER_COLOR = "2F5496"       # Dark Blue

    # Font settings
    HEADER_FONT = "Calibri"
    HEADER_FONT_SIZE = 14
    BODY_FONT = "Calibri"
    BODY_FONT_SIZE = 11

    # Number formats
    CURRENCY_FORMAT = "$#,##0.00"
    PERCENTAGE_FORMAT = "0.00%"
    DATE_FORMAT = "mm/dd/yyyy"
    NUMBER_FORMAT = "#,##0"

    # Chart settings
    CHART_STYLE = 10
    CHART_WIDTH = 15
    CHART_HEIGHT = 10

    # Report settings
    FREEZE_PANES = True
    AUTO_FILTER = True
    COLUMN_WIDTH_DEFAULT = 15
    ROW_HEIGHT_HEADER = 30

    @staticmethod
    def get_report_title(report_type):
        """Generate report title with timestamp"""
        timestamp = datetime.now().strftime("%Y-%m-%d")
        return f"{Config.COMPANY_NAME} - {report_type} - {timestamp}"
