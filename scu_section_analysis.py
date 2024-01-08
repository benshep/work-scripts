import os
import cv2
import numpy as np
import scipy.optimize
from math import pi
from random import randint

folder = os.path.join(os.environ['UserProfile'], 'Documents', 'SCU', '2022-08 Sectioned former')
os.chdir(folder)
for image_number in (
        # 1,
        3, 6, 7, 9, 11, 13, 16, 18, 20, 22):
    # Load image, grayscale, Gaussian blur, Otsu's threshold
    image = cv2.imread(f'{image_number}_cr.jpg')
    height, width, _ = image.shape

    hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
    blur = cv2.GaussianBlur(image, (1, 1), 0)
    red = blur[:, :, 0]
    green = blur[:, :, 1]
    blue = blur[:, :, 2]
    hue = hsv[:, :, 0]
    saturation = hsv[:, :, 1]
    value = hsv[:, :, 2]

    thresh = (blue >= 120) & (blue <= 180) & (hue >= 90) & (hue <= 120)
    # thresh = (value >= 120) & (blue <= 200)
    thresh = thresh.astype(np.uint8) * 255
    # thresh = cv2.Canny(thresh, 100, 200)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
    cv2.imwrite(f'{image_number}_thresh.jpg', thresh)

    # Dilate with elliptical shaped kernel
    # dilate = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)
    # dilate = cv2.morphologyEx(dilate, cv2.MORPH_ERODE, kernel, iterations=2)
    # cv2.imwrite(f'{image_number}_dilate.jpg', dilate)
    # dilate = cv2.dilate(thresh, kernel, iterations=2)

    # Find contours, filter using contour threshold area, draw ellipse
    contours = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = contours[0] if len(contours) == 2 else contours[1]
    ellipses = []
    for c in contours:
        colour = (randint(0, 255), randint(0, 255), randint(0, 255))
        # cv2.drawContours(image, [c], -1, colour, 4)
        area = cv2.contourArea(c)
        if len(c) >= 5:  # and 12000 > area > 1000:
            ellipse = cv2.fitEllipse(c)
            (x, y), (a, b), theta = ellipse
            ellipse_area = pi * a * b
            if 29000 < ellipse_area < 40000 and a / b > 0.7:
                # print(f'x, y = ({x:.0f}, {y:.0f}); (a, b) = ({a:.0f}, {b:.0f}), {theta=:.0f}, {ellipse_area=:.0f}')
                print(x, y, a, b, theta)
                colour = (255, 255, 255)  # (36, 255, 12)
                cv2.ellipse(image, ellipse, colour, 4)
                # 'erase' from thresh image (with a fat marker!) to try to reduce overlap of new fitted ellipses below
                cv2.ellipse(thresh, ellipse, (0, 0, 0), 20)
                label_text = f'{ellipse_area / 1000:.0f}'
                # cv2.putText(image, label_text, (int(x - len(label_text) * 10), int(y)), cv2.FONT_HERSHEY_SIMPLEX, 1, colour, 2)
                ellipses.append([x, y, a, b, theta])
    print(f'{len(ellipses)} ellipses found and labelled.')
    cv2.imwrite(f'{image_number}_thresh.jpg', thresh)
    ellipses = np.array(ellipses)
    print(repr(ellipses))

    # How to find 'holes' in the grid, where not all ellipses are found by this algorithm?
    # Assign some row and column numbers to each ellipse
    avg_ab = np.mean(ellipses[:, 2:4].flat)  # average of a and b ellipse params
    # bigger gap in X direction though
    distance = avg_ab * np.array([1, 0.9])
    min_xy = np.percentile(ellipses[:, 0:2], 5, axis=0)
    round_xy = np.round((ellipses[:, 0:2] - min_xy) / distance).astype(int)
    # remove duplicates
    unique_xy, indices = np.unique(round_xy, axis=0, return_index=True)
    ellipses = ellipses[indices]

    for xy, (x, y, a, b, theta) in zip(unique_xy, ellipses):
        cv2.ellipse(image, ((x, y), (a, b), theta), (0, 255, 255), 2)
        label_text = f'{xy}'
        cv2.putText(image, label_text, (int(x - len(label_text) * 10), int(y)), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2)

    # find holes
    x_min, y_min = np.min(round_xy, axis=0)
    x_max, y_max = np.max(round_xy, axis=0)
    print(f'Rows: {y_max - y_min + 1}; columns: {x_max - x_min + 1}')

    new_ellipses = []
    lhs = np.min(ellipses[:, 0])
    rhs = np.max(ellipses[:, 0])
    for y in np.unique(round_xy[:, 1]):
        find_row = unique_xy[:, 1] == y
        this_row = unique_xy[find_row]
        row_total = len(this_row)
        row_xlist = ellipses[find_row, 0]
        diffs = np.diff(row_xlist)
        avg_gap = np.median(diffs)  # assume more wire than hole!
        hole_locs = (diffs / avg_gap > 0.7).nonzero()[0]
        poly_y = np.polyfit(ellipses[find_row, 0], ellipses[find_row, 1], deg=1)
        poly_a = np.polyfit(ellipses[find_row, 0], ellipses[find_row, 2], deg=1)
        poly_b = np.polyfit(ellipses[find_row, 0], ellipses[find_row, 3], deg=1)
        poly_theta = np.polyfit(ellipses[find_row, 0], ellipses[find_row, 4], deg=1)
        for hole in hole_locs:
            # how big is gap?
            gap = diffs[hole]
            n_missing = np.ceil(0.8 * gap / avg_gap).astype(int) - 1
            row_total += n_missing
            # print(y, hole, n_missing)
            # best guess for hole position(s)
            x_pos = row_xlist[hole] + (np.arange(n_missing) + 1) * (1 / (n_missing + 1)) * diffs[hole]
            # revise estimate of gap (useful for adding more wires outside the row later)
            diffs = np.diff(np.sort(np.append(row_xlist, x_pos)))
            avg_gap = np.median(diffs)
            for xx in x_pos:
                y_pos = np.polyval(poly_y, xx)
                a = np.polyval(poly_a, xx)
                b = np.polyval(poly_b, xx)
                theta = np.polyval(poly_theta, xx)
                new_ellipses.append([xx, y_pos, a, b, theta])
                cv2.ellipse(image, ((xx, y_pos), (a, b), theta), (255, 255, 0), 2)
        needed = 8 + (y - y_min) % 2  # 8 for even, 9 for odd
        full_row = row_total == needed
        left_of_row = row_xlist[0]
        right_of_row = row_xlist[-1]
        while row_total < needed:
            gap_lhs = left_of_row - avg_gap - lhs
            gap_rhs = rhs - right_of_row - avg_gap
            if gap_lhs > gap_rhs:
                left_of_row -= avg_gap
                xx = left_of_row
            else:
                right_of_row += avg_gap
                xx = right_of_row
            row_total += 1
            y_pos = np.polyval(poly_y, xx)
            a = np.polyval(poly_a, xx)
            b = np.polyval(poly_b, xx)
            theta = np.polyval(poly_theta, xx)
            new_ellipses.append([xx, y_pos, a, b, theta])
            try:
                cv2.ellipse(image, ((xx, y_pos), (a, b), theta), (255, 128, 0), 2)
            except cv2.error:
                pass
            print(f'{y}, {gap_lhs:.0f}, {gap_rhs:.0f}')

    cv2.imwrite(f'{image_number}_image.jpg', image)


    def overlap(args):
        xx, yy, aa, bb, th = args
        build_image = np.zeros((height, width, 3), np.uint8)
        ellipse_params = ((xx, yy), (a, b), th)
        # print(ellipse_params)
        try:
            cv2.ellipse(build_image, ellipse_params, (255, 255, 255), 6)
            and_image = build_image[:, :, 0] & thresh
            # cv2.imwrite(f'{image_number}_build.jpg', build_image)
            cv2.imwrite(f'{image_number}_and.jpg', and_image)
            result = sum(and_image.flat)
        except cv2.error:
            result = 0
        # print([int(arg) for arg in args], result)
        return float(-result)


    for start_vals in new_ellipses:
        res = scipy.optimize.minimize(overlap, start_vals, method='Nelder-Mead', tol=10 , options={'eps': 3})
        x, y, a, b, theta = res.x
        ellipses = np.append(ellipses, res.x)
        ellipse_params = ((x, y), (a, b), theta)
        ellipse_area = pi * a * b
        print(f'Revised to ({x:.0f}, {y:.0f}); (a, b) = ({a:.0f}, {b:.0f}), {theta=:.0f}, {ellipse_area=:.0f}')
        cv2.ellipse(image, ellipse_params, (255, 0, 0), 4)
        cv2.imwrite(f'{image_number}_image.jpg', image)

    np.savetxt(f'{image_number}.csv', ellipses, delimiter=',')
