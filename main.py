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

def sun_rise(lat, lon, day, tz, horizon=-18):
    """To get the sun rise for today, we need to ask ephem for the previous rising of the sun starting with midnight tomorrow in this timezone."""

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
    """To get the sun set for today, we need to ask ephem the next setting of the sun starting with midnight today in this timezone."""

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
    """To get the moon rise for today, we need to ask ephem for the previous rising of the moon starting with midnight tomorrow in this timezone."""

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
    """To get the moon set for today, we need to ask ephem the next setting of the moon starting with midnight today in this timezone."""

    # Make an observer
    obs = ephem.Observer()
    obs.lat = str(lat)
    obs.lon = str(lon)
    obs.horizon = str(horizon)

    midnight_today = datetime.datetime.combine(day, datetime.datetime.min.time(),
                                                 tzinfo=tz)
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
            if today.moon_set < today.moon_rise: # moon looks like |---___---|
                # need dusk to be between moon rise and set
                if today.moon_set < today.dusk < today.moon_rise:
                    self.dark_time = True
                    self.start = today.dusk
                    self.end = today.moon_rise
            else: # moon looks like |___---___|
                # need dusk to not be between moon rise and set
                if not (today.moon_set < today.dusk < today.moon_rise):
                    self.dark_time = True
                    self.start = today.dusk
                    self.end = min(tomorrow.dawn, tomorrow.moon_rise)

        if self.dark_time:
            self.duration = self.end - self.start


    def __str__(self):
        if self.dark_time:
            start_str = self.start.strftime("%I:%M %p %Z")

            end_str = self.end.strftime("%I:%M %p %Z   ")
            if self.end.date() != self.start.date():
                end_str = self.end.strftime("%I:%M %p %Z +1")

            duration_str = fmt_timedelta(self.duration).ljust(8, ' ')
        else:
            start_str = ' -          '
            end_str = ' -             '
            duration_str = ' -      '

        return f'| {self.day.strftime("%m/%d/%y")} | {start_str} | {end_str} | {duration_str} |'

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

    # Construct Raw Data

    a_days = {}
    for dt in rrule(DAILY, dtstart=args.start-timedelta(days=1), until=args.end+timedelta(days=1)):
        day = dt.date()
        a_days[day] = AstroDay(args.lat, args.lon, day, args.tz)

    # Construct Dark Times

    dark_times = {}
    for dt in rrule(DAILY, dtstart=args.start, until=args.end):
        today = dt.date()
        yesterday = today - timedelta(days=1)
        tomorrow = today + timedelta(days=1)

        dark_times[today] = DarkTime(a_days[yesterday], a_days[today], a_days[tomorrow])




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


if __name__ == '__main__':
    main(sys.argv[1:])
