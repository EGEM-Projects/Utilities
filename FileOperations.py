import pandas as pd
import logging


from Misc import highlight_headers, highlight_columns, ColoredFormatter, createHeatMap

class ExcelFileHandler:
    """
    The ExcelFileHandler class provides a convenient way to manage and manipulate Excel files using Pandas and openpyxl/xlsxwriter libraries.
    It offers functionalities for reading, writing, and formatting Excel sheets, specifically for data-driven tasks.
    """

    def __init__(self, b_enable_logging: bool, filepath=None):
        """
        Initialize the Excel file handler.

        :param filepath: Path to the Excel file (optional).
        """
        self.filepath = filepath
        self.workbook = None
        self.formats = {}

        # Create a logger
        self.logger = logging.getLogger(self.__class__.__name__)
        if b_enable_logging:
            self.logger.setLevel(logging.DEBUG)
            handler = logging.StreamHandler()
            formatter = ColoredFormatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
        else:
            self.logger.setLevel(logging.ERROR)

        self.logger.info("Initializing ExcelFileHandler class")

    def file_exists(self):
        """Check if the file exists."""
        try:
            open(self.filepath, 'r').close()
            return True
        except FileNotFoundError:
            return False

    def set_filepath(self, filepath):
        """
        Set or update the filepath for the Excel handler.

        :param filepath: Path to the Excel file.
        """
        self.filepath = filepath

    def read_excel(self):
        """
        Read the Excel file.
        """
        return pd.read_excel(self.filepath, engine='openpyxl')


    def read_csv(self):
        """
        Read the Excel file.
        """
        return pd.read_csv(self.filepath)

    def read_sheet(self, sheetname):
        """
        Read the contents of a sheet into a Pandas DataFrame.

        :param sheetname: Name of the sheet to read.
        :return: DataFrame containing the sheet data.
        """
        return pd.read_excel(self.filepath, sheet_name=sheetname)

    def amend_records(self, sheetname, index_id, rows_to_add):
        """
        Amend records in a sheet by removing rows with a specific index_id and adding new rows.

        :param sheetname: Name of the sheet to amend.
        :param index_id: Value to search for in the first column to remove.
        :param rows_to_add: DataFrame of rows to add.
        """
        RecordsAmended = False
        df = self.read_sheet(sheetname)

        # Remove rows with matching index_id
        df = df[df.iloc[:, 0] != index_id]  # Assuming the first column is the index column

        # Add new rows
        df = pd.concat([df, rows_to_add], ignore_index=True)

        # Write back to the Excel file
        with pd.ExcelWriter(self.filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheetname, index=False)
        RecordsAmended = True
        return RecordsAmended

    def delete_rows(self, sheetname, index_id):
        """
        Delete rows where the first column matches index_id.

        :param sheetname: Name of the sheet to modify.
        :param index_id: Value to search for in the first column to remove.
        """
        RowsDeleted = False

        df = self.read_sheet(sheetname)
        df = df[df.iloc[:, 0] != index_id]
        self.write_data(sheetname, df, False)

        RowsDeleted = True
        return RowsDeleted


    def write_data(self, dataframes_dict: dict, index: bool):
        """
        Write multiple Pandas DataFrames to an Excel workbook with dynamic sheet names.

        :param dataframes_dict: Dictionary where keys are sheet names and values are DataFrames.
        :param index: Whether to write DataFrame index.
        :param max_length: Maximum column width (default is 20). If None, it auto-adjusts based on content.
        """
        DataWritten = False

        with pd.ExcelWriter(self.filepath, engine='openpyxl', mode='w') as writer:
            for sheetname, dataframe in dataframes_dict.items():
                dataframe.to_excel(writer, sheet_name=sheetname, index=index)

        DataWritten = True
        return DataWritten



    def write_with_formatting(self, sheets_data, formats):
        """
        Write data to multiple sheets and apply conditional formatting.

        :param sheets_data: List of tuples (sheetname, dataframe, startrow, startcol).
        :param formats: List of tuples for conditional formatting.
        """
        DataFormatted = False

        with pd.ExcelWriter(self.filepath, engine='xlsxwriter') as writer:

            workbook = writer.book

            fmt1 = workbook.add_format({'num_format': '#,##0.00', 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})
            fmt2 = workbook.add_format({'num_format': '#,##0,', 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})

            # Write data to sheets
            for sheetname, dataframe, startrow, startcol in sheets_data:
                if dataframe is not None:  # Ensure the dataframe is not None
                    dataframe.to_excel(writer, sheet_name=sheetname, startrow=startrow, startcol=startcol)

            # Apply conditional formatting
            for sheetname, start_row, start_col, end_row, end_col, fmt, cols_to_highlight in formats:
                worksheet = writer.sheets[sheetname]
                format_to_apply = fmt1 if fmt['format'] == 'fmt' else fmt2
                worksheet.conditional_format(start_row, start_col, end_row, end_col, {'type': fmt['type'], 'format': format_to_apply})
                HeadersHighlighted = highlight_headers(workbook, worksheet, start_row, start_col, end_col) if fmt['colHeaders'] == True else None
                if not HeadersHighlighted and fmt['colHeaders'] == True:
                    self.logger.error(f"Headers have not been highlighted on {sheetname}.")
                ColumnsHighlighted = highlight_columns(workbook, worksheet, start_row, end_row, cols_to_highlight) if fmt['colsToHighlight'] == True else None
                if not ColumnsHighlighted and fmt['colsToHighlight'] == True:
                    self.logger.error(f"Column have not been highlighted on sheet {sheetname}.")
                HeatMapCreated = createHeatMap(worksheet, start_row, start_col, end_row, end_col) if fmt['createHeatMap'] == True else None
                if not HeatMapCreated and fmt['createHeatMap'] == True:
                    self.logger.error(f"HeatMap failed to create on sheet {sheetname}.")
                # Autofit Worksheet
                worksheet.autofit()

        DataFormatted = True
        return DataFormatted















