import os
import numpy as np
import xml.etree.ElementTree as ET
import matplotlib.pyplot as plt

use_raw_data = False
os.chdir(r"\\ddast0029\d\ZEPTO\1mm (2)")
prev_int = None
grads = []
for x in range(-10, 11):
    # x = 0
    filename = f"{x}.xml"
    try:
        root = ET.parse(filename).getroot()
    except FileNotFoundError:
        continue
    doubles = root.findall('Cluster/DBL')
    params = {double.find('Name').text: float(double.find('Val').text) for double in doubles}

    if use_raw_data:
        # get raw data
        arrays = root.findall('Array')
        raw_data = next(array for array in arrays if array.find('Name').text == 'Raw Data')
        dims = tuple(int(dim.text) for dim in raw_data.findall('Dimsize'))

        doubles = raw_data.findall('DBL')
        values = np.array([float(double.find('Val').text) for double in doubles]).reshape(dims)
        voltage = values[0]  # first all the voltage values are listed
        # print(*voltage, sep='\n')
        dt = values[1]  # then all the time differences
        assert all(dt == dt[0])  # they should all be the same
        dt = dt[0]

        # where is the maximum?
        # max_i = np.argmin(voltage) if abs(min(voltage)) > max(voltage) else np.argmax(voltage)
        if x == 0:
            max_i = 8180
        else:
            total_integral = np.sum(voltage)
            compare = np.less if total_integral < 0 else np.greater
            # where does the integral cross half the total integral? (midpoint of S-curve)
            max_i = np.where(compare(np.cumsum(voltage), total_integral / 2))[0][0]
        # assume baseline is 1 second before this
        before_i = max(0, max_i - round(1 / dt))
        # and 2 seconds after
        n_avgs = 200  # should be a multiple of 20 due to mains noise at 50Hz (assuming dt = 1ms)
        after_i = min(len(voltage) - n_avgs, max_i + round(1 / dt))
        before_avg = np.average(voltage[before_i:(before_i + n_avgs)])
        after_avg = np.average(voltage[after_i:(after_i + n_avgs)])
        slope = (after_avg - before_avg) / (after_i - before_i)
        int_start = before_i + n_avgs
        int_end = after_i
        baseline_v = np.linspace(before_avg + slope * n_avgs / 2,
                                 after_avg - slope * n_avgs / 2, int_end - int_start)
        # baseline_v = np.average(np.hstack([voltage[before_i:(before_i + n_avgs)],
        #                                    voltage[after_i:(after_i + n_avgs)]]))
        corrected_voltage = voltage[int_start:int_end] - baseline_v
        integral = np.sum(corrected_voltage)
        # print(int_end - int_start)
        if x == 0:
            plt.plot(np.cumsum(corrected_voltage))
            # plt.plot(voltage)  # [int_start:int_end])
            plt.title(x)
            plt.show(block=True)
        # print(*voltage[int_start:int_end], sep='\t')
        grad = integral - prev_int if prev_int else 0
        prev_int = integral
        grads.append(grad)
    print(
        # x or 1e-99,  # 0 -> 1e-99 since Excel is funny about polynomial functions of zero
        params['dY (v.s)'] * -1e6,
        # integral * 1e6,
        # grad * 1e6,
        sep='\t')

# print(np.average(grads), np.std(grads))
