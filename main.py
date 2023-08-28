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
import tzlocal
from datetime import timedelta


def main():
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

        # Terminal printout to help with troubleshooting
        print('Begin astronomical twilight:', begin_twilight)
        print('End astronomical twilight:', end_twilight)
        print('Moonrise time:', moonrise_time)
        print('Moonset time:', moonset_time)
        print()

        # Increment the current date by one day
        current_utc += 1

    # Create Excel file using xlsxwriter
    workbook = xlsxwriter.Workbook('darkSkyTimes.xlsx')
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
        # Included %Y/%m/%d to help with troubleshooting
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
        # Included %Y/%m/%d to help with troubleshooting
        moon_rise_str = moon_rise.strftime('%Y/%m/%d %H:%M')
        local_moon_rise_day = moon_rise.astimezone(pytz.timezone('US/Mountain')).day + 1
        moon_rise_month = moon_rise.astimezone(pytz.timezone('US/Mountain')).month
        moon_rise_year = moon_rise.astimezone(pytz.timezone('US/Mountain')).year
        first, last = monthrange(moon_rise_year, moon_rise_month)
        if local_moon_rise_day > last:
            local_moon_rise_day = local_moon_rise_day - last

        # Included %Y/%m/%d to help with troubleshooting
        end_twilight_sheet = end_twilight_times[row].replace(tzinfo=pytz.timezone('UTC')).astimezone(
            pytz.timezone('US/Mountain')).strftime('%Y/%m/%d %H:%M')
        # Included %Y/%m/%d to help with troubleshooting
        begin_twilight_sheet = begin_twilight_times[row].replace(tzinfo=pytz.timezone('UTC')).astimezone(
            pytz.timezone('US/Mountain')).strftime('%Y/%m/%d %H:%M')

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


class AstroDay:
    def __str__(self):
        return f'The day is: {self.day}, {self.moon_rise = }, {self.moon_set = }, ' \
               f'{self.twilight_evening = }, {self.twilight_morning = }'

    def __init__(self, day):
        self.day = day
        self.moon_rise = None
        self.moon_set = None
        self.twilight_evening = None
        self.twilight_morning = None

    def populate_astro_data(self, obs):
        # To get the moon set and twilight evening for today, we need to ask ephem the next setting of the
        # moon and sun starting with midnight today in this timezone.
        midnight_tonight = datetime.datetime.combine(self.day, datetime.datetime.min.time(),
                                                     tzinfo=tzlocal.get_localzone())
        e_date = ephem.Date(midnight_tonight)
        obs.date = e_date
        obs.horizon = '0'
        moon_set_unaware_dt = obs.next_setting(ephem.Moon()).datetime()
        obs.horizon = '-18'
        twilight_evening_unaware_dt = obs.next_setting(ephem.Sun(), use_center=True).datetime()

        # To get the moon rise and twilight morning for today, we need to ask ephem for the previous rising
        # of the moon and sun starting with midnight tomorrow in this timezone.
        midnight_tomorrow = datetime.datetime.combine(self.day + datetime.timedelta(days=1),
                                                      datetime.datetime.min.time(),
                                                      tzinfo=tzlocal.get_localzone())
        e_date = ephem.Date(midnight_tomorrow)
        obs.date = e_date
        obs.horizon = '0'
        moon_rise_unaware_dt = obs.previous_rising(ephem.Moon()).datetime()
        obs.horizon = '-18'
        twilight_morning_unaware_dt = obs.previous_rising(ephem.Sun(), use_center=True).datetime()

        self.moon_set = pytz.utc.localize(moon_set_unaware_dt)
        self.moon_rise = pytz.utc.localize(moon_rise_unaware_dt)
        self.twilight_morning = pytz.utc.localize(twilight_morning_unaware_dt)
        self.twilight_evening = pytz.utc.localize(twilight_evening_unaware_dt)


def test():
    today = datetime.date(year=2023, month=5, day=8)
    next_day = today + timedelta(days=1)
    current_day = today.day
    next_day_number = next_day.day
    d = AstroDay(today)
    d_next = AstroDay(next_day)
    print(d)
    print(d_next)

    # Make an observer
    obs = ephem.Observer()
    lat = '40.7720'
    lon = '-112.1012'
    obs.lat = lat
    obs.lon = lon

    d.populate_astro_data(obs)
    print(d)
    print(d.moon_rise.astimezone())
    print(d.moon_set.astimezone())
    print(d.twilight_morning.astimezone())
    print(d.twilight_evening.astimezone())
    print()

    d_next.populate_astro_data(obs)
    print(d_next)
    print(d_next.moon_rise.astimezone())
    print(d_next.moon_set.astimezone())
    print(d_next.twilight_morning.astimezone())
    print(d_next.twilight_evening.astimezone())
    print()

    moon_rise_tz = d.moon_rise.astimezone()
    moon_set_tz = d.moon_set.astimezone()
    twilight_morning_tz = d.twilight_morning.astimezone()
    twilight_evening_tz = d.twilight_evening.astimezone()

    next_moon_rise_tz = d_next.moon_rise.astimezone()
    next_moon_set_tz = d_next.moon_set.astimezone()
    next_twilight_morning_tz = d_next.twilight_morning.astimezone()
    next_twilight_evening_tz = d_next.twilight_evening.astimezone()

    moon_rise_day = moon_rise_tz.day
    moon_set_day = moon_set_tz.day
    twilight_morning_day = twilight_morning_tz.day
    twilight_evening_day = twilight_evening_tz.day

    next_moon_rise_day = next_moon_rise_tz.day
    next_moon_set_day = next_moon_set_tz.day
    next_twilight_morning_day = next_twilight_morning_tz.day
    next_twilight_evening_day = next_twilight_evening_tz.day

    if moon_set_day == current_day:
        print(moon_set_tz)
    else:
        no_moon_set = 'null'
        print(no_moon_set)

    print(twilight_evening_tz)
    print(twilight_morning_tz)

    if moon_rise_day == current_day:
        print(moon_rise_tz)
    else:
        no_moon_rise = 'null'
        print(no_moon_rise)
    print()

    # Get times for next day
    if next_moon_set_day == next_day_number:
        print(next_moon_set_tz)
    else:
        no_moon_set = 'null'
        print(no_moon_set)

    print(next_twilight_evening_tz)
    print(next_twilight_morning_tz)

    if next_moon_rise_day == next_day_number:
        print(next_moon_rise_tz)
    else:
        no_moon_rise = 'null'
        print(no_moon_rise)
    print()

    # There are 5 conditions that produce dark sky.
    # Condition 1: From moon set to twilight morning on same day.
    #               When the moon sets before twilight morning on the same day.
    if (moon_set_day == current_day) and (moon_set_tz < twilight_morning_tz):
        set_morning_duration = twilight_morning_tz - moon_set_tz
        print(set_morning_duration)
    # Condition 2: From twilight evening to moon rise on the same day.
    #               When the moon rise happens after twilight evening on the same day.
    elif (moon_rise_day == current_day) and (twilight_evening_tz < moon_rise_tz):
        evening_rise_duration = moon_rise_tz - twilight_evening_tz
        print(evening_rise_duration)
    # Condition 3: From moon set to twilight morning of the next day.
    # ***OFFSET NEED TO FIX
    #                when the moon sets after twilight evening on one day
    #                and rises after twilight morning on the next day.
    elif (moon_set_day == current_day) and (twilight_evening_tz < moon_set_tz) and \
            (next_twilight_morning_tz < next_moon_rise_tz):
        set_next_morning_duration = next_twilight_morning_tz - moon_set_tz
        print(set_next_morning_duration)
    # Condition 4: From twilight evening to twilight morning on the next day
    # ***OFFSET NEED TO FIX
    #               When the moon sets before twilight evening
    #               and does not rise before twilight evening on the same day
    elif (moon_set_day == current_day) and (moon_set_tz < twilight_evening_tz) and \
            (moon_rise_tz < twilight_evening_tz) and (next_twilight_morning_tz < next_moon_rise_tz):
        evening_next_morning_duration = next_twilight_morning_tz - twilight_evening_tz
        print(evening_next_morning_duration)
    # Condition 5: From twilight evening to moon rise of the next day
    #               When the moon rise for the next day happens before twilight morning on the next day.
    elif (moon_rise_day == current_day) and (moon_set_tz < twilight_evening_tz) and \
            (next_moon_rise_tz < next_twilight_morning_tz):
        evening_next_rise_duration = next_moon_rise_tz - twilight_evening_tz
        print(evening_next_rise_duration)
    else:
        print("0")

    # NEED TO FIGURE OUT DURATION OFFSET
    # bug when there is no moon rise for condition 5



if __name__ == '__main__':
    #    main()
    test()

# TODO
# Throw out or set to none when ephem returns a date time that isn't today.
# Loop through all the dates in our range.
