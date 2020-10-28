from datetime import datetime, timedelta
import pytz
import matplotlib.pyplot as plt
from ics import Calendar
from urllib.request import urlopen
import xlwings as xw


class Shift:
    def __init__(self, start_date, end_date):
        self.date = start_date.date()
        self.day = self.date.strftime("%A")
        self.start_time = start_date.time()
        self.end_time = end_date.time()
        self.hours = get_worked_hours(start_date, end_date)
        self.p_rate_pay = get_penalty_rate_pay(start_date, end_date)
        self.norm_rate_pay = round(self.hours * wage, 2)
        self.pay = round(self.norm_rate_pay + self.p_rate_pay, 2)

    def __repr__(self):
        return "[Date: {}\n Day: {}\n Start: {}\n End: {}\n Hours Worked: {}\n Pay: ${}\n" \
               " Normal Rate Pay: ${}\n Penalty Rate Pay: ${}]".format(
                self.date, self.day, self.start_time, self.end_time, self.hours, self.pay, self.norm_rate_pay,
                self.p_rate_pay)

    def __add__(self, other):
        if self.date == other.date:
            if self.start_time > other.start_time:
                start_time = self.start_time
                end_time = other.end_time
            else:
                start_time = self.start_time
                end_time = other.end_time

            start = datetime.combine(date=self.date, time=start_time)
            end = datetime.combine(date=self.date, time=end_time)

            return Shift(start, end)

        else:
            return None


def get_worked_hours(start_time, end_time):
    delta = end_time - start_time
    hours = float(delta.seconds / 3600)
    # Subtract breaks
    if hours > 7:
        hours -= 1
    elif hours > 5:
        hours -= 0.5

    return hours


def get_penalty_rate_pay(start_time, end_time):
    # Weekday
    if start_time.weekday() < 5:
        if end_time.hour >= 18:
            worked = end_time.hour - 18
            return worked * late_rate
        return 0
    # Saturday
    elif start_time.weekday() == 5:
        return get_worked_hours(start_time, end_time) * sat_rate
    # Sunday
    elif start_time.weekday() == 6:
        early_hours = 0
        worked = 0
        if start_time.hour < 9:
            early_hours = (9 - start_time.hour)
            worked = early_hours * early_rate

        pay = ((get_worked_hours(start_time, end_time) - early_hours) * sun_rate + worked)

        return round(pay, 2)

    return 0


def to_date(t):
    date = datetime.strptime(str(t)[:-9].replace('T', ' '), '%Y-%m-%d %H%M')
    time = pytz.timezone("UTC").localize(date).astimezone(pytz.timezone("Australia/Sydney"))
    date = datetime(time.year, time.month, time.day, time.hour, time.minute)
    return date


def calculate_pay_for_week(l):
    pay = 0

    for shift in l:
        pay += shift.pay

    sheet.range('A5').value = pay
    tax = sheet.range('B5').value
    pay -= tax

    return round(pay, 2), tax


def get_week_range(date):
    weekday = date.weekday()
    start_of_week = datetime(date.year, date.month, date.day-weekday)
    end_of_week = start_of_week + timedelta(days=7)
    print(start_of_week)
    print(end_of_week)
    return start_of_week, end_of_week


def add_shift(shift, l):
    for s in l:
        if s.date == shift.date:
            combined_shift = s + shift
            l.append(combined_shift)
            l.remove(s)
            l.sort(key=lambda x: x.date)
            return
    l.append(shift)
    l.sort(key=lambda x: x.date)


def get_week_events():
    url = 'url.ics'
    c = Calendar(urlopen(url).read().decode('iso-8859-1'))
    print(c.events[-1])
    a = to_date(c.events[-1].begin)
    b = to_date(c.events[-1].end)
    s = Shift(a,b)
    print(s)
    for event in c.events:
        a = to_date(event.begin)
        b = to_date(event.end)
        if dt_from <= a <= dt_to:
            add_shift(Shift(a, b), shifts)
            shifts.sort(key=lambda shift: shift.date)


wage = 14.74
sat_rate = wage * .25
late_rate = wage * .25
sun_rate = wage * .8
early_rate = wage * 2.0

# https://www.ato.gov.au/Rates/Weekly-tax-table/
xl = xw.Book('tax.xlsx')
sheet = xl.sheets[0]
excel = xw.apps.active
shifts = []


dt_from, dt_to = get_week_range(datetime.now())
get_week_events()

weekly_pay, weekly_tax = calculate_pay_for_week(shifts)
summary = '\${}, \${} in tax'.format(weekly_pay, weekly_tax)

print(shifts)
print(summary)

# PLOTTING #

def get_plot_lists():
    pos, value, name, lab = ([None] * 7 for z in range(4))
    for i in range(7):
        for shift in shifts:
            if shift.date.weekday() == i:
                name[i] = shift.date.strftime('%A')
                value[i] = shift.pay
                lab[i] = '${}'.format(shift.pay)
                break
        if value[i] is None:
            new_date = dt_from + timedelta(days=i)
            name[i] = (new_date.strftime('%A'))
            value[i] = 0
            lab[i] = '$0'
        pos[i] = i+1

    return pos, value, name, lab


positions, height, label, labels = get_plot_lists()

fig, ax = plt.subplots(num=None, figsize=(10, 6), dpi=80, facecolor='w', edgecolor ='k')

ax.bar(positions, height, tick_label=label, width=.8, color=['green'])
ax.set_ylabel("Pay in $")
ax.set_xlabel("Days")
ax.set_title('Pay for week of {}\n '.format(str(dt_from.date())) + summary)
for i in range(len(positions)):
    ax.text(x=positions[i]-.2, y=height[i]+1, s=labels[i], size=10)

fig.canvas.set_window_title('Pay for week')

plt.show()

