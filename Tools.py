from datetime import date


def number_of_months(d1, d2):
    return (d1.year - d2.year) * 12 + d1.month - d2.month


def cubic_spline(input_column, output_column, x):
    input_count = len(input_column)
    output_count = len(output_column)

    if input_count != output_count:
        return "The number of indices and the number of output_columns don't match"

    xin = input_column
    yin = output_column

    n = input_count
    yt = [0] * n
    u = [0] * (n - 1)

    for i in range(1, n - 1):
        sig = (xin[i] - xin[i - 1]) / (xin[i + 1] - xin[i - 1])
        p = sig * yt[i - 1] + 2
        yt[i] = (sig - 1) / p
        u[i] = (yin[i + 1] - yin[i]) / (xin[i + 1] - xin[i]) - (yin[i] - yin[i - 1]) / (xin[i] - xin[i - 1])
        u[i] = (6 * u[i] / (xin[i + 1] - xin[i - 1]) - sig * u[i - 1]) / p

    qn = 0
    un = 0
    yt[n - 1] = (un - qn * u[n - 2]) / (qn * yt[n - 2] + 1)

    for k in range(n - 2, -1, -1):
        yt[k] = yt[k] * yt[k + 1] + u[k]

    klo = 0
    khi = n - 1
    while khi - klo > 1:
        k = (khi - klo) // 2
        if xin[k] > x:
            khi = k
        else:
            klo = k

    h = xin[khi] - xin[klo]
    a = (xin[khi] - x) / h
    b = (x - xin[klo]) / h
    y = a * yin[klo] + b * yin[khi] + ((a ** 3 - a) * yt[klo] + (b ** 3 - b) * yt[khi]) * (h ** 2) / 6

    return y


def linear_interp(x_arr, y_arr, x):
    if x < x_arr[0] or x > x_arr[-1]:
        raise ValueError(f"linear Interpolation: x is out of bound. Lower bound= {x_arr[0]} and Upper bound= {x_arr[-1]}")

    if x_arr[0] == x:
        return y_arr[0]

    for i in range(len(x_arr)):
        if x_arr[i] >= x:
            return y_arr[i - 1] + (x - x_arr[i - 1]) / (x_arr[i] - x_arr[i - 1]) * (y_arr[i] - y_arr[i - 1])
