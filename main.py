import argparse
import sys
import xlsxwriter
import ephem
import pytz
import datetime
import tzlocal

from datetime import datetime as dt
from datetime import timedelta
from calendar import monthrange
from openpyxl import load_workbook
from openpyxl.styles import numbers
from dateutil import parser, tz
from dateutil.rrule import rrule, DAILY
from dateutil.relativedelta import relativedelta

def sun_rise(lat, lon, date, tz, horizon=-18):
    # Make an observer
    obs = ephem.Observer()
    obs.lat = str(lat)
    obs.lon = str(lon)
    obs.horizon = str(horizon)

    return None

def sun_set(lat, lon, date, tz, horizon=-18):
    # Make an observer
    obs = ephem.Observer()
    obs.lat = str(lat)
    obs.lon = str(lon)
    obs.horizon = str(horizon)

    return None

def moon_rise(lat, lon, date, tz, horizon=0):
    # Make an observer
    obs = ephem.Observer()
    obs.lat = str(lat)
    obs.lon = str(lon)
    obs.horizon = str(horizon)

    return None

def moon_set(lat, lon, day, tz, horizon=0):
    # Make an observer
    obs = ephem.Observer()
    obs.lat = str(lat)
    obs.lon = str(lon)
    obs.horizon = str(horizon)

    midnight_tonight = datetime.datetime.combine(day, datetime.datetime.min.time(),
                                                 tzinfo=tz)
    obs.date = ephem.Date(midnight_tonight)
    moon_set_unaware_dt = pytz.utc.localize(obs.next_setting(ephem.Moon()).datetime())

    moon_set = moon_set_unaware_dt.astimezone(tz)

    # print(f'{day = }, {str(moon_set_unaware_dt) = } {str(moon_set) = }')

    if moon_set.date() != day:
        return None
    return moon_set

class AstroDay:
    def __str__(self):
        dt_fmt = "%I:%M %p %Z"

        if self.dawn:
            dawn_str = self.dawn.strftime(dt_fmt)
        else:
            dawn_str = ' -          '

        if self.dusk:
            dusk_str = self.moon_rise.strftime(dt_fmt)
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


    def __init__(self, lat, lon, day, tz):
        self.lat = lat
        self.lon = lon
        self.day = day
        self.tz = tz

        self.moon_rise = moon_rise(self.lat, self.lon, self.day, self.tz)
        self.moon_set = moon_set(self.lat, self.lon, self.day, self.tz)
        self.dawn = sun_rise(self.lat, self.lon, self.day, self.tz)
        self.dusk = sun_set(self.lat, self.lon, self.day, self.tz)

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

def parse_args(args):
    parser = argparse.ArgumentParser(description='Generate audio file for a Ms. Reddit Youtube video')
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

    print('+----------+--------------+--------------+--------------+--------------+')
    print('| Date     | Dawn         | Dusk         | Moon Rise    | Moon Set     |')
    print('+----------+--------------+--------------+--------------+--------------+')

    for day in rrule(DAILY, dtstart=args.start, until=args.end):
        print(AstroDay(args.lat, args.lon, day.date(), args.tz))

    print('+----------+--------------+--------------+--------------+--------------+')

if __name__ == '__main__':
    main(sys.argv[1:])
