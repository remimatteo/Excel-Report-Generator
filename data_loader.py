"""
Data loading module - supports multiple data sources
"""
import pandas as pd
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

class DataLoader:
    """Load data from various sources"""

    @staticmethod
    def load_csv(file_path, **kwargs):
        """
        Load data from CSV file

        Args:
            file_path: Path to CSV file
            **kwargs: Additional pandas read_csv arguments

        Returns:
            pd.DataFrame: Loaded data
        """
        try:
            logger.info(f"Loading data from CSV: {file_path}")
            df = pd.read_csv(file_path, **kwargs)
            logger.info(f"Loaded {len(df)} rows, {len(df.columns)} columns")
            return df
        except Exception as e:
            logger.error(f"Error loading CSV: {str(e)}")
            raise

    @staticmethod
    def load_excel(file_path, sheet_name=0, **kwargs):
        """
        Load data from Excel file

        Args:
            file_path: Path to Excel file
            sheet_name: Sheet name or index
            **kwargs: Additional pandas read_excel arguments

        Returns:
            pd.DataFrame: Loaded data
        """
        try:
            logger.info(f"Loading data from Excel: {file_path}, sheet: {sheet_name}")
            df = pd.read_excel(file_path, sheet_name=sheet_name, **kwargs)
            logger.info(f"Loaded {len(df)} rows, {len(df.columns)} columns")
            return df
        except Exception as e:
            logger.error(f"Error loading Excel: {str(e)}")
            raise

    @staticmethod
    def load_json(file_path, **kwargs):
        """
        Load data from JSON file

        Args:
            file_path: Path to JSON file
            **kwargs: Additional pandas read_json arguments

        Returns:
            pd.DataFrame: Loaded data
        """
        try:
            logger.info(f"Loading data from JSON: {file_path}")
            df = pd.read_json(file_path, **kwargs)
            logger.info(f"Loaded {len(df)} rows, {len(df.columns)} columns")
            return df
        except Exception as e:
            logger.error(f"Error loading JSON: {str(e)}")
            raise

    @staticmethod
    def create_sample_data():
        """
        Create sample sales data for demonstration

        Returns:
            pd.DataFrame: Sample data
        """
        import numpy as np
        from datetime import datetime, timedelta

        # Generate sample data
        np.random.seed(42)
        dates = pd.date_range(start='2024-01-01', end='2024-12-31', freq='D')

        data = {
            'Date': dates,
            'Region': np.random.choice(['North', 'South', 'East', 'West'], size=len(dates)),
            'Product': np.random.choice(['Product A', 'Product B', 'Product C', 'Product D'], size=len(dates)),
            'Sales': np.random.randint(1000, 10000, size=len(dates)),
            'Quantity': np.random.randint(10, 100, size=len(dates)),
            'Cost': np.random.randint(500, 5000, size=len(dates)),
        }

        df = pd.DataFrame(data)

        # Calculate profit
        df['Profit'] = df['Sales'] - df['Cost']
        df['Profit_Margin'] = (df['Profit'] / df['Sales'] * 100).round(2)

        logger.info(f"Created sample data with {len(df)} rows")
        return df

    @staticmethod
    def validate_data(df, required_columns=None):
        """
        Validate loaded data

        Args:
            df: DataFrame to validate
            required_columns: List of required column names

        Returns:
            bool: True if valid, raises exception otherwise
        """
        if df is None or df.empty:
            raise ValueError("DataFrame is empty")

        if required_columns:
            missing = set(required_columns) - set(df.columns)
            if missing:
                raise ValueError(f"Missing required columns: {', '.join(missing)}")

        # Check for completely empty columns
        empty_cols = df.columns[df.isnull().all()].tolist()
        if empty_cols:
            logger.warning(f"Found completely empty columns: {', '.join(empty_cols)}")

        logger.info("Data validation passed")
        return True
