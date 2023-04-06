from datetime import date

def year_fraction(dt1, dt2, dt3, cnt, freq):
    if cnt == "30/360":
        return dc_30_360(dt1, dt2, freq)
    elif cnt == "ACT/ACT":
        return act_act(dt1, dt2, dt3, freq)
    elif cnt == "ACT/360":
        return act_360(dt1, dt2, freq)
    elif cnt == "ACT/365":
        return act_365(dt1, dt2, freq)
    else:
        print("Not correct day count input")
        return None

def dc_30_360(dt1, dt2, freq):
    d1, d2 = dt1.day, dt2.day
    m1, m2 = dt1.month, dt2.month
    Y1, Y2 = dt1.year, dt2.year

    if d1 == 31:
        d1 = 30
    if d2 == 31 and d1 >= 30:
        d2 = 30
    if m1 == 2 and dt1.year % 4 == 0 and d1 == 29:
        d1 = 30
    if m1 == 2 and dt1.year % 4 != 0 and d1 == 28:
        d1 = 30

    return -(((360 * (Y2 - Y1)) + (30 * (m2 - m1)) + (d2 - d1)) / 360) * freq

def act_360(dt1, dt2, freq):
    return -(dt1 - dt2).days / 360 * freq * 360 / 366

def act_365(dt1, dt2, freq):
    return -(dt1 - dt2).days / 366 * freq

def act_act(value_date, start_date, nxt_date, freq):
    nxt_date = start_date.replace(month=start_date.month + 12 // freq)

    if (nxt_date - start_date).days == 0:
        return 0
    else:
        return (nxt_date - value_date).days / ((nxt_date - start_date).days)


"""
To use these functions in Python,
you can create date objects using the date(year, month, day) constructor from the datetime module. For example:

date1 = date(2023, 1, 1)
date2 = date(2023, 6, 30)
date3 = date(2023, 7, 1)
cnt = "30/360"
freq = 2

result = year_fraction(date1, date2, date3, cnt, freq)
print(result)
"""