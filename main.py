from dateutil import parser
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
import xlsxwriter
import ephem
import pytz
import datetime

ROW_OFFSET = 1

# Make an observer
obs = ephem.Observer()

#  lat = input("Enter your latitude: ")
#  lon = input("Enter your longitude: ")
lat = '40.7720'
lon = '-112.1012'
start_date = parser.parse(input("Enter the start date (YYYY/MM/DD): "))
end_date = parser.parse(input("Enter the end date (YYYY/MM/DD): "))

obs.lat = lat
obs.lon = lon
start_utc = ephem.Date(start_date)
end_utc = ephem.Date(end_date + relativedelta(days=1))

current_utc = start_utc

moonset_times = []
end_twilight_times = []
begin_twilight_times = []
moonrise_times = []

while current_utc < end_utc:

    obs.date = current_utc

    # Calculate the moonrise and moonset times
    obs.horizon = '0'
    moonset_time = obs.next_setting(ephem.Moon(), start=current_utc)

    moonrise_time = obs.previous_rising(ephem.Moon(), start=current_utc)

    # Calculate the astronomical twilight times
    obs.horizon = '-18'  # astronomical twilight
    begin_twilight = obs.previous_rising(ephem.Sun(), start=current_utc)
    end_twilight = obs.next_setting(ephem.Sun(), start=current_utc)

    # Convert the UTC times to the observer's local timezone
    obs.date = begin_twilight
    begin_twilight_local = ephem.localtime(obs.date).strftime('%Y/%m/%d %I:%M %p')
    obs.date = end_twilight
    end_twilight_local = ephem.localtime(obs.date).strftime('%Y/%m/%d %I:%M %p')
    obs.date = moonrise_time
    moonrise_time_local = ephem.localtime(obs.date).strftime('%Y/%m/%d %I:%M %p')
    obs.date = moonset_time
    moonset_time_local = ephem.localtime(obs.date).strftime('%Y/%m/%d %I:%M %p')

    moonset_times.append(moonset_time.datetime())  # add the local time zone moonset time to the list
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
    workbook = xlsxwriter.Workbook("sheetTest3.xlsx")

    bold_format = workbook.add_format({'bold': True})

    # format cells
    cell_format = workbook.add_format()
    cell_format.set_text_wrap()
    cell_format.set_align('center_across')

    time_format = workbook.add_format({'num_format': 'h:mm AM/PM'})
    green_time_format = workbook.add_format({'num_format': 'h:mm AM/PM'})
    green_time_format.set_font_color('green')

    worksheet = workbook.add_worksheet('DarkSkyTimes')

    #  create headers using cell location, then what you want to name it
    worksheet.write('A2', 'Day', bold_format)
    worksheet.write('B2', 'Moon Set', bold_format)
    worksheet.write('C2', 'Astro. Twilight END', bold_format)
    worksheet.write('D2', 'Astro. Twilight START', bold_format)
    worksheet.write('E2', 'Moon Rise', bold_format)
    worksheet.write('F2', 'Duration', bold_format)

    rowIndex = 3

    #  for row in range(31):
    for row, moon_set_sheet in enumerate(moonset_times):
        day_sheet = row + ROW_OFFSET
        # if you want to verify that the day matches the row
        moon_set_tz = moon_set_sheet.replace(tzinfo=pytz.timezone('UTC'))
        moon_set = moon_set_tz.astimezone(pytz.timezone('US/Mountain'))
        moon_set_str = moon_set.strftime('%Y/%m/%d %H:%M %p')
        moon_day = moon_set.astimezone(pytz.timezone('US/Mountain')).day
        if moon_set.astimezone(pytz.timezone('US/Mountain')).day != day_sheet:
            moon_set_str = ' '
        moon_rise_tz = moonrise_times[row].replace(tzinfo=pytz.timezone('UTC'))
        moon_rise = moon_rise_tz.astimezone(pytz.timezone('US/Mountain'))
        moon_rise_str = moon_rise.strftime('%Y/%m/%d %H:%M %p')
        if moon_rise.astimezone(pytz.timezone('US/Mountain')).day != day_sheet:
            moon_rise_str = ' '

        end_twilight_sheet = end_twilight_times[row].replace(tzinfo=pytz.timezone('UTC')).astimezone(pytz.timezone('US/Mountain')).strftime('%Y/%m/%d %H:%M %p')
        begin_twilight_sheet = begin_twilight_times[row].replace(tzinfo=pytz.timezone('UTC')).astimezone(pytz.timezone('US/Mountain')).strftime('%Y/%m/%d %H:%M %p')
        duration_sheet = row + 7

        worksheet.write('A' + str(rowIndex + 1), day_sheet, cell_format)
        worksheet.write('B' + str(rowIndex + 1), moon_set_str, cell_format)
        worksheet.write('C' + str(rowIndex), end_twilight_sheet, cell_format)
        worksheet.write('D' + str(rowIndex), begin_twilight_sheet, cell_format)
        worksheet.write('E' + str(rowIndex), moon_rise_str, cell_format)
        worksheet.write('F' + str(rowIndex), duration_sheet, cell_format)

        rowIndex += 1

        # print(day_sheet, moon_set_sheet, end_twilight_sheet, start_twilight_sheet, moon_rise_sheet, duration_sheet)

    #  Set column size. Numbering starts at 0, so to resize column A do 0,0
    worksheet.set_column(0, 0, width=3)
    worksheet.set_column(1, 1, width=20)
    worksheet.set_column(2, 2, width=20)
    worksheet.set_column(3, 3, width=20)
    worksheet.set_column(4, 4, width=20)
    worksheet.set_column(5, 5, width=7)
    workbook.close()


if __name__ == "__main__":
    main()

