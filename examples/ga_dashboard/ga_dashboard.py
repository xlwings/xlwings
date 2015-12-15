# Follow the blog post here: http://blog.zoomeranalytics.com/google-analytics/

import os
import numpy as np
import pandas as pd
import pandas.io.ga as ga
from datetime import datetime
from xlwings import Workbook, Range

# Worksheets
sheet_dashboard = 'Sheet1'

# Client Secrets file: same dir as this file
client_secrets = os.path.abspath(os.path.join(os.path.dirname(__file__),
                                              'client_secrets.json'))


def behavior(start_date, end_date, account_name, property_name, profile_name, max_results):
    """
    Writes a DataFrame with the number of pageviews per half-hours x weekdays
    to the Range "behavior"
    """
    # Let pandas fetch the data from Google Analytics, returns a generator object
    df_chunks = ga.read_ga(secrets=client_secrets,
                           account_name=account_name,
                           property_name=property_name,
                           profile_name=profile_name,
                           dimensions=['date', 'hour', 'minute'],
                           metrics=['pageviews'],
                           start_date=start_date,
                           end_date=end_date,
                           index_col=0,
                           parse_dates={'datetime': ['date', 'hour', 'minute']},
                           date_parser=lambda x: datetime.strptime(x, '%Y%m%d %H %M'),
                           max_results=max_results,
                           chunksize=10000)

    # Concatenate the chunks into a DataFrame and get number of rows
    df = pd.concat(df_chunks)
    num_rows = df.shape[0]

    # Resample into half-hour buckets
    df = df.resample('30Min', how='sum')

    # Create the behavior table (half-hour x weekday)
    grouped = df.groupby([df.index.time, df.index.weekday])
    behavior = grouped['pageviews'].aggregate(np.sum).unstack()

    # Make sure the table covers all hours and weekdays
    behavior = behavior.reindex(index=pd.date_range("00:00", "23:30", freq="30min").time,
                                columns=range(7))
    behavior.columns = ['MO', 'TU', 'WE', 'TH', 'FR', 'SA', 'SU']

    # Write to Excel.
    # Time-only values are currently a bit of a pain on Windows, so we set index=False.
    Range(sheet_dashboard, 'behavior', index=False).value = behavior
    Range(sheet_dashboard, 'effective').value = num_rows


def refresh():
    """
    Refreshes the tables in Excel given the input parameters.
    """
    # Connect to the Workbook
    wb = Workbook.caller()

    # Read input
    start_date = Range(sheet_dashboard, 'start_date').value
    end_date = Range(sheet_dashboard, 'end_date').value
    account_name = Range(sheet_dashboard, 'account').value
    property_name = Range(sheet_dashboard, 'property').value
    profile_name = Range(sheet_dashboard, 'view').value
    max_results = Range(sheet_dashboard, 'max_rows').value

    # Clear Content
    Range(sheet_dashboard, 'behavior').clear_contents()
    Range(sheet_dashboard, 'effective').clear_contents()

    # Behavior table
    behavior(start_date, end_date, account_name, property_name, profile_name, max_results)


if __name__ == '__main__':
    # To run from Python. Not needed when called from Excel.
    path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'ga_dashboard.xlsm'))
    Workbook.set_mock_caller(path)
    refresh()