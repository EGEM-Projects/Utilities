import os
import logging
import win32com.client


class ExcelMacroRunner:
    """
    ExcelMacroRunner allows you to execute macros in Excel (.xlsm) files
    using the Windows COM interface (via win32com.client).
    """

    def __init__(self, filepath: str, enable_logging: bool = True):
        """
        Initialize the macro runner.

        :param filepath: Path to the .xlsm Excel file.
        :param enable_logging: Enable or disable logging.
        """
        self.filepath = os.path.abspath(filepath)
        self.enable_logging = enable_logging
        self.logger = logging.getLogger(self.__class__.__name__)

        if enable_logging:
            self.logger.setLevel(logging.DEBUG)
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            if not self.logger.hasHandlers():
                self.logger.addHandler(handler)
        else:
            self.logger.disabled = True

        if not os.path.exists(self.filepath):
            raise FileNotFoundError(f"Excel file not found: {self.filepath}")

    def run_macro(self, macro_name: str, visible: bool = False, save: bool = True) -> bool:
        """
        Run a macro from the Excel workbook.

        :param macro_name: Name of the macro (e.g., 'MyMacro').
        :param visible: Whether to show Excel while running the macro.
        :param save: Whether to save the workbook after macro execution.
        :return: True if successful, False otherwise.
        """
        try:
            self.logger.info("Launching Excel via COM interface...")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = visible

            self.logger.info(f"Opening workbook: {self.filepath}")
            workbook = excel.Workbooks.Open(self.filepath)

            full_macro_name = f"'{os.path.basename(self.filepath)}'!{macro_name}"
            self.logger.info(f"Running macro: {full_macro_name}")
            excel.Application.Run(full_macro_name)

            if save:
                workbook.Save()
                self.logger.info("Workbook saved after macro run.")

            workbook.Close(False)
            excel.Quit()
            self.logger.info("Excel application closed.")
            return True

        except Exception as e:
            self.logger.error(f"Error running macro '{macro_name}': {e}")
            return False
