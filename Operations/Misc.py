import logging
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill


class ColoredFormatter(logging.Formatter):
    # Define colors for different log levels
    COLORS = {
        "DEBUG": "\033[37m",  # White
        "INFO": "\033[32m",   # Green
        "WARNING": "\033[33m",  # Yellow
        "ERROR": "\033[31m",   # Red
        "CRITICAL": "\033[1;31m",  # Bold Red
    }
    RESET = "\033[0m"  # Reset color

    def format(self, record):
        log_color = self.COLORS.get(record.levelname, self.RESET)
        message = super().format(record)
        return f"{log_color}{message}{self.RESET}"


def get_highlight_style(value, style_type, highlight_styles):
    """
    Get the style for a given value and style type.

    :param value: The value to be styled.
    :param style_type: The type of style ('pnl', 'stoploss', 'var').
    :return: The inline style string for the cell.
    """
    for lower, upper, color in highlight_styles[style_type]:
        if lower <= value <= upper:
            return f'background-color:{color}'
    return 'background-color:white'  # Default


def highlight_negative(value):
    """Returns red color for negative values in HTML formatting"""
    color = "red" if isinstance(value, (int, float)) and value < 0 else "black"
    return f"color: {color};"

def highlight_bold(value):
    font_type = "bold" if(isinstance(value, str)) else "font-style:bold"
    return f"font-weight: {font_type};"

def highlight_italic(value):
    font_type = "italic" if(isinstance(value, str)) else "font-style:italic"
    return f"font-style: {font_type};"

def highlight_headers(workbook, worksheet, start_row, start_col, end_col):

    HeadersHighlighted = False
    # Define header format
    header_format = workbook.add_format({"bold":True,"text_wrap":False,"valign":"top","bg_color":"#00008B","font_color":"white","border":1})
    # Apply the format to each header cell
    worksheet.conditional_format(start_row, start_col, start_row, end_col, {
            'type': 'no_blanks',  # Cell-based condition
            'format': header_format  # Format to apply
        }
    )
    HeadersHighlighted = True
    return HeadersHighlighted

def highlight_columns(workbook, worksheet, start_row, end_row, cols_to_highlight):

    ColumnsHighlighted = False
    highlight_format = workbook.add_format({"bg_color": "#FFEB9C"})  # Light yellow fill
    for index in cols_to_highlight:
        worksheet.conditional_format(start_row, index + 1, end_row, index + 1, {"type": "no_blanks", "format": highlight_format})

    ColumnsHighlighted = True
    return ColumnsHighlighted

def createHeatMap(worksheet, start_row, start_col, end_row, end_col):

    CreateHeatMap = False
    worksheet.conditional_format(start_row, start_col, end_row, end_col, {'type': '3_color_scale',
        'min_color': "#FF0000",  # Blue
        'mid_color': "#FFFFFF",  # White
        'max_color': "#0000FF"  # Red
    })
    CreateHeatMap = True
    return CreateHeatMap


def weighted_avg(df, values, weights):
    d = df[values]
    w = df[weights]
    weighted_sum = (d * w).sum()
    total_weight = w.sum()

    return weighted_sum / total_weight if total_weight != 0 else 0  # Avoid division by zero