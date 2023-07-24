from dateutil import parser
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
import xlsxwriter
import ephem
import pytz
import datetime
from calendar import monthrange
from openpyxl import load_workbook
from openpyxl.styles import numbers

ROW_OFFSET = 1

# Make an observer
obs = ephem.Observer()

"""
Disabled lat/lon and date input prompt so you don't have to type them in every time while troubleshooting
lat = input("Enter your latitude: ")
lon = input("Enter your longitude: ")
start_date = parser.parse(input("Enter the start date (YYYY/MM/DD): "))
end_date = parser.parse(input("Enter the end date (YYYY/MM/DD): "))
"""
lat = '40.7720'
lon = '-112.1012'
start_date = parser.parse('2023/05/01')
end_date = parser.parse('2023/06/01')

obs.lat = lat
obs.lon = lon
start_utc = ephem.Date(start_date)
end_utc = ephem.Date(end_date + relativedelta(days=1))

current_utc = start_utc

moonset_times = []
end_twilight_times = []
begin_twilight_times = []
moonrise_times = []

# Note: The Ephem library ALWAYS uses Universal Time (UTC), never the local time zone, which complicates things
while current_utc < end_utc:

    obs.date = current_utc

    # Calculate the moonrise and moonset times
    obs.horizon = '0'
    moonset_time = obs.next_setting(ephem.Moon(), start=current_utc)
    moonrise_time = obs.previous_rising(ephem.Moon(), start=current_utc)

    # Calculate the astronomical twilight times (when sun is -18 below horizon)
    obs.horizon = '-18'
    begin_twilight = obs.previous_rising(ephem.Sun(), start=current_utc, use_center=True)
    end_twilight = obs.next_setting(ephem.Sun(), start=current_utc, use_center=True)

    # Convert the UTC times to the observer's local timezone
    obs.date = begin_twilight
    begin_twilight_local = ephem.localtime(obs.date).strftime('%Y/%m/%d %I:%M %p')
    obs.date = end_twilight
    end_twilight_local = ephem.localtime(obs.date).strftime('%Y/%m/%d %I:%M %p')
    obs.date = moonrise_time
    moonrise_time_local = ephem.localtime(obs.date).strftime('%Y/%m/%d %I:%M %p')
    obs.date = moonset_time
    moonset_time_local = ephem.localtime(obs.date).strftime('%Y/%m/%d %I:%M %p')

    # Add the local time zone moonset/rise and astronomical twilight begin/end times to the list
    moonset_times.append(moonset_time.datetime())
    end_twilight_times.append(end_twilight.datetime())
    begin_twilight_times.append(begin_twilight.datetime())
    moonrise_times.append(moonrise_time.datetime())

    print('Begin astronomical twilight:', begin_twilight)
    print('End astronomical twilight:', end_twilight)
    print('Moonrise time:', moonrise_time)
    print('Moonset time:', moonset_time)
    print()

    # Increment the current date by one day
    current_utc += 1


def main():
    # Create Excel file using xlsxwriter
    workbook = xlsxwriter.Workbook("darkSkyTimes.xlsx")
    worksheet = workbook.add_worksheet('DarkSkyTimes')

    # Format cells
    cell_header_format = workbook.add_format()
    cell_header_format.set_bold()
    cell_header_format.set_text_wrap()
    cell_header_format.set_align('center_across')

    time_format = workbook.add_format()
    time_format.set_text_wrap()
    time_format.set_align('center_across')

    # Create headers using cell location
    worksheet.write('A2', 'Day', cell_header_format)
    worksheet.write('B2', 'Moon Set', cell_header_format)
    worksheet.write('C2', 'Astronomical Twilight END', cell_header_format)
    worksheet.write('D2', 'Astronomical Twilight START', cell_header_format)
    worksheet.write('E2', 'Moon Rise', cell_header_format)
    worksheet.write('F2', 'Duration', cell_header_format)

    rowIndex = 3

    for row, moon_set_sheet in enumerate(moonset_times):
        day_sheet = row + ROW_OFFSET
        moon_set_tz = moon_set_sheet.replace(tzinfo=pytz.timezone('UTC'))
        moon_set = moon_set_tz.astimezone(pytz.timezone('US/Mountain'))
        moon_set_str = moon_set.strftime('%Y/%m/%d %H:%M')
        local_moon_set_day = moon_set.astimezone(pytz.timezone('US/Mountain')).day

        """
        On days when the moon does not rise or does not set in that day, then that cell should be blank.
        However, the rise/set time for the next day was being populated in that cell that should be blank.
        This solves that problem, but creates a new problem where the last day or two of the month is
        skipping a day or being displayed on the wrong day.

        Current problem:
        Second to last day of month: incorrect moon rise time, it displays the moon rise time for the next day.
        Last day of month: incorrect moon set time, it displays the time for the next day (which is the first day
        of the next month). Also, no data is being filled in for the last day of the month for the moon rise time,
        because it is being inputted on the second to last day of the month, so that cell is blank when it shouldn't be.
        """
        moon_rise_tz = moonrise_times[row].replace(tzinfo=pytz.timezone('UTC'))
        moon_rise = moon_rise_tz.astimezone(pytz.timezone('US/Mountain'))
        moon_rise_str = moon_rise.strftime('%Y/%m/%d %H:%M')
        local_moon_rise_day = moon_rise.astimezone(pytz.timezone('US/Mountain')).day + 1
        moon_rise_month = moon_rise.astimezone(pytz.timezone('US/Mountain')).month
        moon_rise_year = moon_rise.astimezone(pytz.timezone('US/Mountain')).year
        first, last = monthrange(moon_rise_year, moon_rise_month)
        if local_moon_rise_day > last:
            local_moon_rise_day = local_moon_rise_day - last

        end_twilight_sheet = end_twilight_times[row].replace(tzinfo=pytz.timezone('UTC')).astimezone(pytz.timezone('US/Mountain')).strftime('%Y/%m/%d %H:%M')
        begin_twilight_sheet = begin_twilight_times[row].replace(tzinfo=pytz.timezone('UTC')).astimezone(pytz.timezone('US/Mountain')).strftime('%Y/%m/%d %H:%M')

        worksheet.write('A' + str(rowIndex + 1), day_sheet, cell_header_format)
        if local_moon_set_day != day_sheet:
            worksheet.write('B' + str(rowIndex), moon_set_str, time_format)
        else:
            worksheet.write('B' + str(rowIndex + 1), moon_set_str, time_format)
        worksheet.write('C' + str(rowIndex), end_twilight_sheet, time_format)
        worksheet.write('D' + str(rowIndex), begin_twilight_sheet, time_format)
        if local_moon_rise_day != day_sheet:
            worksheet.write('E' + str(rowIndex - 1), moon_rise_str, time_format)
        else:
            worksheet.write('E' + str(rowIndex), moon_rise_str, time_format)

        rowIndex += 1


    # Set column size
    worksheet.set_column('A:A', 5)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 23)
    worksheet.set_column('D:D', 25)
    worksheet.set_column('E:E', 21)
    worksheet.set_column('F:F', 9)

    workbook.close()

    # Load the workbook using openpyxl
    workbook = load_workbook('darkSkyTimes.xlsx')

    # Get the active sheet
    sheet = workbook.active

    # Delete the row of times that is generated before the specified date-range since Ephem uses "previous_rising"
    sheet.delete_rows(3)

    # Save the modified workbook
    workbook.save('darkSkyTimes.xlsx')


if __name__ == "__main__":
    main()

