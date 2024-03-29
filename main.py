import argparse
import sys
import ephem
import pytz
import datetime
import tzlocal
import pandas as pd

from datetime import timedelta
from dateutil import parser, tz
from dateutil.rrule import rrule, DAILY


def sun_rise(lat, lon, day, tz, horizon=-18):
    """
    To get the sun rise for today, we need to ask ephem for the previous
    rising of the sun starting with midnight tomorrow in this timezone.
    """

    # Make an observer
    obs = ephem.Observer()
    obs.lat = str(lat)
    obs.lon = str(lon)
    obs.horizon = str(horizon)

    midnight_tomorrow = datetime.datetime.combine(day + datetime.timedelta(days=1),
                                                  datetime.datetime.min.time(),
                                                  tzinfo=tzlocal.get_localzone())
    obs.date = ephem.Date(midnight_tomorrow)
    sun_rise_unaware_dt = obs.previous_rising(ephem.Sun(), use_center=True).datetime()
    sun_rise = pytz.utc.localize(sun_rise_unaware_dt).astimezone(tz)

    # print(f'{day = }, {str(sun_rise_unaware_dt) = } {str(sun_rise) = }')

    if sun_rise.date() != day:
        return None
    return sun_rise


def sun_set(lat, lon, day, tz, horizon=-18):
    """
    To get the sun set for today, we need to ask ephem the next
    setting of the sun starting with midnight today in this timezone.
    """

    # Make an observer
    obs = ephem.Observer()
    obs.lat = str(lat)
    obs.lon = str(lon)
    obs.horizon = str(horizon)

    midnight_today = datetime.datetime.combine(day, datetime.datetime.min.time(),
                                               tzinfo=tz)
    obs.date = ephem.Date(midnight_today)

    sun_set_unaware_dt = obs.next_setting(ephem.Sun(), use_center=True).datetime()
    sun_set = pytz.utc.localize(sun_set_unaware_dt).astimezone(tz)

    # print(f'{day = }, {str(sun_set_unaware_dt) = } {str(sun_set) = }')

    if sun_set.date() != day:
        return None
    return sun_set


def moon_rise(lat, lon, day, tz, horizon=0):
    """
    To get the moon rise for today, we need to ask ephem for the previous
    rising of the moon starting with midnight tomorrow in this timezone.
    """

    # Make an observer
    obs = ephem.Observer()
    obs.lat = str(lat)
    obs.lon = str(lon)
    obs.horizon = str(horizon)

    midnight_tomorrow = datetime.datetime.combine(day + datetime.timedelta(days=1),
                                                  datetime.datetime.min.time(),
                                                  tzinfo=tzlocal.get_localzone())
    obs.date = ephem.Date(midnight_tomorrow)

    moon_rise_unaware_dt = obs.previous_rising(ephem.Moon()).datetime()
    moon_rise = pytz.utc.localize(moon_rise_unaware_dt).astimezone(tz)

    # print(f'{day = }, {str(moon_rise_unaware_dt) = } {str(moon_rise) = }')

    if moon_rise.date() != day:
        return None
    return moon_rise


def moon_set(lat, lon, day, tz, horizon=0):
    """
    To get the moon set for today, we need to ask ephem the next
    setting of the moon starting with midnight today in this timezone.
    """

    # Make an observer
    obs = ephem.Observer()
    obs.lat = str(lat)
    obs.lon = str(lon)
    obs.horizon = str(horizon)

    midnight_today = datetime.datetime.combine(day, datetime.datetime.min.time(), tzinfo=tz)
    obs.date = ephem.Date(midnight_today)

    moon_set_unaware_dt = obs.next_setting(ephem.Moon()).datetime()
    moon_set = pytz.utc.localize(moon_set_unaware_dt).astimezone(tz)

    # print(f'{day = }, {str(moon_set_unaware_dt) = } {str(moon_set) = }')

    if moon_set.date() != day:
        return None
    return moon_set


def fmt_timedelta(td):
    hour, rem = divmod(td.seconds, 3600)
    min, sec = divmod(rem, 60)
    return f'{hour}h {min}m'


class AstroDay:
    def __init__(self, lat, lon, day, tz):
        self.lat = lat
        self.lon = lon
        self.day = day
        self.tz = tz

        self.moon_rise = moon_rise(self.lat, self.lon, self.day, self.tz)
        self.moon_set = moon_set(self.lat, self.lon, self.day, self.tz)
        self.dawn = sun_rise(self.lat, self.lon, self.day, self.tz)
        self.dusk = sun_set(self.lat, self.lon, self.day, self.tz)

    def better_format(self):
        dt_fmt = "%I:%M %p %Z"

        if self.dawn:
            dawn_str = self.dawn.strftime(dt_fmt)
        else:
            dawn_str = ' -          '

        if self.dusk:
            dusk_str = self.dusk.strftime(dt_fmt)
        else:
            dusk_str = ' -          '

        if self.moon_rise:
            moon_rise_str = self.moon_rise.strftime(dt_fmt)
        else:
            moon_rise_str = ' -          '

        if self.moon_set:
            moon_set_str = self.moon_set.strftime(dt_fmt)
        else:
            moon_set_str = ' -          '

        return [self.day.strftime("%m/%d/%y"), dawn_str, dusk_str, moon_rise_str, moon_set_str]

    def __str__(self):
        dt_fmt = "%I:%M %p %Z"

        if self.dawn:
            dawn_str = self.dawn.strftime(dt_fmt)
        else:
            dawn_str = ' -          '

        if self.dusk:
            dusk_str = self.dusk.strftime(dt_fmt)
        else:
            dusk_str = ' -          '

        if self.moon_rise:
            moon_rise_str = self.moon_rise.strftime(dt_fmt)
        else:
            moon_rise_str = ' -          '

        if self.moon_set:
            moon_set_str = self.moon_set.strftime(dt_fmt)
        else:
            moon_set_str = ' -          '

        return f'| {self.day.strftime("%m/%d/%y")} | {dawn_str} | {dusk_str} | {moon_rise_str} | {moon_set_str} |'


# this should really be two 0 / 1 waveforms in utc, that's the right datastructure, and just poll as needed
# dark times would then just be whenever both of them are 0
class DarkTime:
    def __init__(
            self,
            yesterday: AstroDay,
            today: AstroDay,
            tomorrow: AstroDay
    ):
        self.day = today.day
        self.dark_time = False

        # moon sets while sun is down
        if today.moon_set:
            if today.moon_set < today.dawn:
                self.dark_time = True
                self.start = today.moon_set
                self.end = today.dawn
            elif today.moon_set > today.dusk:
                self.dark_time = True
                self.start = today.moon_set
                self.end = min(tomorrow.dawn, tomorrow.moon_rise)

        # sun sets while moon is down
        if not today.moon_set:
            if today.dusk < today.moon_rise:
                self.dark_time = True
                self.start = today.dusk
                self.end = today.moon_rise
            # otherwise false
        elif not today.moon_rise:
            if today.dusk > today.moon_set:
                self.dark_time = True
                self.start = today.dusk
                self.end = min(tomorrow.dawn, tomorrow.moon_rise)
            # otherwise false
        else:
            if today.moon_set < today.moon_rise:  # moon looks like |---___---|
                # need dusk to be between moon rise and set
                if today.moon_set < today.dusk < today.moon_rise:
                    self.dark_time = True
                    self.start = today.dusk
                    self.end = today.moon_rise
            else:  # moon looks like |___---___|
                # need dusk to not be between moon rise and set
                if not today.moon_set > today.dusk:
                    if not (today.moon_set < today.dusk < today.moon_rise):
                        self.dark_time = True
                        self.start = today.dusk
                        self.end = min(tomorrow.dawn, tomorrow.moon_rise)

        if self.dark_time:
            self.duration = self.end - self.start

    def __str__(self):
        if self.dark_time:
            start_str = self.start.strftime("%I:%M %p %Z")

            end_str = self.end.strftime("%I:%M %p %Z")
            if self.end.date() != self.start.date():
                end_str = self.end.strftime("%I:%M %p %Z +1")

            duration_str = fmt_timedelta(self.duration).ljust(8, ' ')
        else:
            start_str = ' -          '
            end_str = ' -             '
            duration_str = ' -      '

        return f'| {self.day.strftime("%m/%d/%y")} | {start_str} | {end_str} | {duration_str} |'

    def better_format(self):
        if self.dark_time:
            start_str = self.start.strftime("%I:%M %p %Z")

            end_str = self.end.strftime("%I:%M %p %Z")
            if self.end.date() != self.start.date():
                end_str = self.end.strftime("%I:%M %p %Z +1")

            duration_str = fmt_timedelta(self.duration).ljust(8, ' ')
        else:
            start_str = ' -          '
            end_str = ' -             '
            duration_str = ' -      '

        return [self.day.strftime("%m/%d/%y"), start_str, end_str, duration_str]


def parse_args(args):
    parser = argparse.ArgumentParser(description='Generate darkSkyTimes Excel sheet')
    parser.add_argument('-lat', type=float, help='Observer Latitude')
    parser.add_argument('-lon', type=float, help='Observer Longitude')
    parser.add_argument('-start', type=lambda s: datetime.datetime.strptime(s, '%m/%d/%Y').date(),
                        help='Start date (dd/mm/yyyy)')
    parser.add_argument('-end', type=lambda s: datetime.datetime.strptime(s, '%m/%d/%Y').date(),
                        help='End date inclusive (dd/mm/yyyy)')
    parser.add_argument('-tz', choices=pytz.all_timezones, help='The timezone')
    return parser.parse_args(args)


def main(args_list):
    args = parse_args(args_list)
    args.tz = tz.gettz(args.tz)

    # Construct Raw Data into dict
    a_days = {}
    a_days_csv = {}
    for dt in rrule(DAILY, dtstart=args.start - timedelta(days=1), until=args.end + timedelta(days=1)):
        day = dt.date()
        astro_day1 = AstroDay(args.lat, args.lon, day, args.tz)
        a_days[day] = astro_day1
        a_days_csv[day] = astro_day1.better_format()

    # Construct Dark Times into dict
    dark_times = {}
    dark_times_csv = {}
    for dt in rrule(DAILY, dtstart=args.start, until=args.end):
        today = dt.date()
        yesterday = today - timedelta(days=1)
        tomorrow = today + timedelta(days=1)
        dark_times[today] = DarkTime(a_days[yesterday], a_days[today], a_days[tomorrow])
        dark_times_csv[today] = DarkTime(a_days[yesterday], a_days[today], a_days[tomorrow]).better_format()

    # Print data to console
    print('+----------+--------------+--------------+--------------+--------------+')
    print('| Date     | Dawn         | Dusk         | Moon Rise    | Moon Set     |')
    print('+----------+--------------+--------------+--------------+--------------+')

    for day, a_day in a_days.items():
        print(a_day)

    print('+----------+--------------+--------------+--------------+--------------+')

    print()

    print('+----------+--------------+-----------------+----------+')
    print('| Date     | Start        | End             | Duration |')
    print('+----------+--------------+-----------------+----------+')

    for day, dark_time in dark_times.items():
        print(dark_time)

    print('+----------+--------------+-----------------+----------+')

    # Use Pandas to make things pretty :)
    # Create DataFrames from the dicts
    df_data = pd.DataFrame.from_dict(a_days_csv, orient='index')
    df_dark = pd.DataFrame.from_dict(dark_times_csv, orient='index')

    # Create an Excel writer object
    writer = pd.ExcelWriter('darkSkyTimes.xlsx', engine='xlsxwriter')

    # Create Excel file with separate sheets for each DataFrame
    df_dark.to_excel(writer, sheet_name='darkTimes', startrow=1, header=False, index=False)
    df_data.to_excel(writer, sheet_name='rawData', startrow=1, header=False, index=False)

    # Get workbook/worksheet objects
    workbook = writer.book
    ws_dark = writer.sheets['darkTimes']
    ws_raw = writer.sheets['rawData']

    # Define formats to use
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center_across',
        'font_size': 15,
        # 'bg_color': '#808080',   //too dark
        'bg_color': '#D4D4D4',
        'border': 1,
    })

    center_format = workbook.add_format({
        'align': 'center_across',
        'font_size': 14,
    })

    # Set the column width for the dark_times_sheet and add the column headers
    # ***There's probably a better way to do this?
    dark_times_sheet = ws_dark
    dark_times_sheet.set_column('A:A', 12, center_format)
    dark_times_sheet.set_column('B:B', 17, center_format)
    dark_times_sheet.set_column('C:C', 25, center_format)
    dark_times_sheet.set_column('D:D', 14, center_format)
    dark_times_sheet.write('A1', 'Date', header_format)
    dark_times_sheet.write('B1', 'Start', header_format)
    dark_times_sheet.write('C1', 'End', header_format)
    dark_times_sheet.write('D1', 'Duration', header_format)

    # Set the column width for the raw_data_sheet and add the column headers
    # ***There's probably a better way to do this?
    raw_data_sheet = ws_raw
    raw_data_sheet.set_column('A:A', 12, center_format)
    raw_data_sheet.set_column('B:E', 17, center_format)
    raw_data_sheet.write('A1', 'Date', header_format)
    raw_data_sheet.write('B1', 'Dawn', header_format)
    raw_data_sheet.write('C1', 'Dusk', header_format)
    raw_data_sheet.write('D1', 'Moon Rise', header_format)
    raw_data_sheet.write('E1', 'Moon Set', header_format)

    # Ctrl + S
    workbook.close()


if __name__ == '__main__':
    main(sys.argv[1:])

    '''
    TO DO:
    --Test various lat/long locations in various timezones to make sure anyone can use this
        --Tested with Atascadero, CA (UTC-8)
            35.4490
            -120.6165
            Appears to work, but displays in MDT
    
    --How to make this easily deployable with simple GUI for user input?
    
    --Nicer Excel formatting:
    Format alternating rows with dark/light colors to help with readability
    
    --Check if inputs are within a permitted range.
    '''