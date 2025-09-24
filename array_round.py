# https://stackoverflow.com/questions/13483430/how-to-make-rounded-percentages-add-up-to-100#answer-34959983
# Fair rounding for numbers in an array that must sum to a particular number

from math import isclose, sqrt
from heapq import nsmallest


def error_gen(actual, rounded):
    divisor = sqrt(1.0 if actual < 1.0 else actual)
    return abs(rounded - actual) ** 2 / divisor


def fair_round(array: list[float], to_nearest: float = 1) -> list[float] | list[int]:
    target_sum = sum(array)
    rounded_down = [to_nearest * (x // to_nearest) for x in array]
    up_count = round((target_sum - sum(rounded_down)) / to_nearest)
    errors = [(error_gen(x, r + to_nearest) - error_gen(x, r), i) for i, (x, r) in enumerate(zip(array, rounded_down))]
    for _, i in nsmallest(up_count, errors):
        rounded_down[i] += to_nearest
    return rounded_down


def round_to_100(percents):
    if not isclose(sum(percents), 100):
        raise ValueError
    n = len(percents)
    rounded = [int(x) for x in percents]
    up_count = 100 - sum(rounded)
    # print(100, sum(rounded), up_count)
    errors = [(error_gen(percents[i], rounded[i] + 1) - error_gen(percents[i], rounded[i]), i) for i in range(n)]
    rank = sorted(errors)
    for i in range(up_count):
        rounded[rank[i][1]] += 1
    return rounded


if __name__ == '__main__':
    for a in ([13.626332, 47.989636, 9.596008, 28.788024],
              [33.3333333, 33.3333333, 33.3333333],
              [24.25, 23.25, 27.25, 25.25],
              [1.25, 2.25, 3.25, 4.25, 89.0]
              ):
        fr = fair_round(a, 1)
        r100 = round_to_100(a)
        print(sum(a), *[f'{x:.1f}' for x in fr], f'{sum(fr):.1f}', r100, sum(r100))
