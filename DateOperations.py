from datetime import datetime, timedelta
import pandas as pd
from typing import List, Optional
from Misc import ColoredFormatter
import logging

class DateOperations:
    """
    The DateOperations class provides functionality to manage dates while considering holidays and weekends, specifically designed for financial or business contexts.
    This class ensures accurate date computations in a business context by accounting for non-working days (weekends and holidays) and provides easy access to adjusted dates in standard and concatenated formats.
    """

    def __init__(self, b_enable_logging : bool,  Region: str, strCurrentDate: str, strPreviousDate: Optional[str] = None):

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

        self.logger.info("Initializing DateOperations class")

        self.__m_HolidayCalenderList = pd.to_datetime(
            pd.read_csv(r'X:\Dept-Market_Risk_LNG\Python Scripts\Pnl Explained\static\holiday_calendar.csv').query(f"Region == '{Region}'")['Date'], format='%d/%m/%Y').to_list()

        strCurrentDate = datetime.strptime(strCurrentDate, '%Y-%m-%d')
        is_weekday = strCurrentDate.weekday() < 5
        if is_weekday and strCurrentDate not in self.__m_HolidayCalenderList:
            self.__m_CurrentDate = strCurrentDate
            self.__m_CurrentDateConcat = self.__m_CurrentDate.strftime('%Y%m%d')
        else:
            raise ValueError(f"Current date {strCurrentDate.strftime('%Y-%m-%d')} is not a weekday or is a holiday.")

        bgetPriorWorkingDate = False
        if strPreviousDate is None:
            self.__m_PriorDate = self.__m_CurrentDate - timedelta(days=1)
            while not bgetPriorWorkingDate:
                if self.__m_PriorDate.weekday() <5 and self.__m_PriorDate not in self.__m_HolidayCalenderList:
                    bgetPriorWorkingDate = True
                else:
                    self.__m_PriorDate -= timedelta(days=1)
            self.__m_PriorDateConcat = self.__m_PriorDate.strftime('%Y%m%d')
        else:
            strPreviousDate = datetime.strptime(strPreviousDate, '%Y-%m-%d')
            self.__m_PriorDate = strPreviousDate
            self.__m_PriorDateConcat = self.__m_PriorDate.strftime('%Y%m%d')

        if strPreviousDate is None:
            # Check if the difference between current date and prior date is one working day
            if (self.__m_CurrentDate - self.__m_PriorDate).days != 1 or self.__m_PriorDate.weekday() >= 5:
                # Adjust for weekends (if the prior date is Friday, the current date must be Monday)
                if not (
                        self.__m_PriorDate.weekday() == 4 and self.__m_CurrentDate.weekday() == 0
                ):
                    raise ValueError(
                        f"The current date {self.__m_CurrentDate.strftime('%Y-%m-%d')} and prior date {self.__m_PriorDate.strftime('%Y-%m-%d')} must have a one-working-day difference."
                    )


    @property
    def m_HolidayCalenderList(self) -> List[datetime]:
        return self.__m_HolidayCalenderList

    @m_HolidayCalenderList.setter
    def m_HolidayCalenderList(self, m_HolidayCalenderList: List[datetime]) -> None:
        self.__m_HolidayCalenderList = m_HolidayCalenderList

    @property
    def m_CurrentDate(self) -> datetime:
        return self.__m_CurrentDate

    @m_CurrentDate.setter
    def m_CurrentDate(self, input_date: str) -> None:
        input_date = datetime.strptime(input_date, '%Y-%m-%d')
        is_weekday = input_date.weekday() < 5
        if is_weekday and input_date not in self.__m_HolidayCalenderList:
            self.__m_CurrentDate = input_date

    @property
    def m_PriorDate(self) -> datetime:
        return self.__m_PriorDate

    @m_PriorDate.setter
    def m_PriorDate(self, input_date: str) -> None:
        input_date = datetime.strptime(input_date, '%Y-%m-%d')
        is_weekday = input_date.weekday() < 5
        if is_weekday and input_date not in self.__m_HolidayCalenderList:
            self.__m_PriorDate = input_date

    @property
    def m_CurrentDateConcat(self) -> str:
        return self.__m_CurrentDateConcat

    @property
    def m_PriorDateConcat(self) -> str:
        return self.__m_PriorDateConcat

    def m_LastBusinessDayPrevMonth(self):
        # Calculate the last day of the previous month
        first_day_of_current_month = self.__m_CurrentDate.replace(day=1)
        last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)

        # Iterate backwards to find the last working day
        while last_day_of_previous_month.weekday() >= 5 or last_day_of_previous_month in self.__m_HolidayCalenderList:
            last_day_of_previous_month -= timedelta(days=1)

        return last_day_of_previous_month

    def m_LastBusinessDayPrevYear(self):
        # Calculate the last day of the previous month
        previous_year = self.__m_CurrentDate.year - 1
        last_day_of_previous_year = datetime(previous_year, 12, 31)

        # Subtract days until the last working day is found
        while last_day_of_previous_year.weekday() >= 5 or last_day_of_previous_year in self.__m_HolidayCalenderList:
            last_day_of_previous_year -= timedelta(days=1)

        return last_day_of_previous_year

    def m_MonthName(self):
        return self.__m_CurrentDate.strftime('%B')

