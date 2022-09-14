import os
import cv2
import numpy as np
import scipy.optimize
from math import pi
from random import randint

# Load image, grayscale, Gaussian blur, Otsu's threshold
folder = os.path.join(os.environ['UserProfile'], 'Documents', 'SCU', '2022-08 Sectioned former')
os.chdir(folder)
image = cv2.imread('1.jpg')[714:1844, 303:1895]
height, width, _ = image.shape

hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
# [print(f'{h},{s},{v}') for h, s, v in hsv[826, 533:1033]]
blur = cv2.GaussianBlur(image, (1, 1), 0)
red = blur[:, :, 0]
green = blur[:, :, 1]
blue = blur[:, :, 2]
hue = hsv[:, :, 0]
saturation = hsv[:, :, 1]
value = hsv[:, :, 2]

thresh = cv2.threshold(blue, 170, 255, cv2.THRESH_TOZERO_INV)[1]  # + cv2.THRESH_OTSU)[1]
thresh = cv2.threshold(thresh, 130, 255, cv2.THRESH_BINARY)[1]  # + cv2.THRESH_OTSU)[1]
thresh = (blue >= 120) & (blue <= 180) & (hue >= 90) & (hue <= 120)
# thresh = (value >= 120) & (blue <= 200)
thresh = thresh.astype(np.uint8) * 255
# thresh = cv2.Canny(thresh, 100, 200)
kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
cv2.imwrite('1_thresh.jpg', thresh)

# Dilate with elliptical shaped kernel
dilate = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)
dilate = cv2.morphologyEx(dilate, cv2.MORPH_ERODE, kernel, iterations=2)
cv2.imwrite('1_dilate.jpg', dilate)
# dilate = cv2.dilate(thresh, kernel, iterations=2)

# Find contours, filter using contour threshold area, draw ellipse
contours = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
contours = contours[0] if len(contours) == 2 else contours[1]
ellipses = []
for i, c in enumerate(contours):
    colour = (randint(0, 255), randint(0, 255), randint(0, 255))
    # cv2.drawContours(image, [c], -1, colour, 4)
    # compute the center of the contour
    # M = cv2.moments(c)
    # cX = M["m10"] / M["m00"]
    # cY = M["m01"] / M["m00"]
    # label_text = f'{i}'
    # cv2.putText(image, label_text, (int(cX - len(label_text) * 10), int(cY)), cv2.FONT_HERSHEY_SIMPLEX, 1, colour, 2)
    area = cv2.contourArea(c)
    if len(c) >= 5:  # and 12000 > area > 1000:
        ellipse = cv2.fitEllipse(c)
        (x, y), (a, b), theta = ellipse
        ellipse_area = pi * a * b
        if 29000 < ellipse_area < 40000 and a / b > 0.7:
            print(f'x, y = ({x:.0f}, {y:.0f}); (a, b) = ({a:.0f}, {b:.0f}), {theta=:.0f}, {ellipse_area=:.0f}')
            colour = (255, 255, 255)  # (36, 255, 12)
            cv2.ellipse(image, ellipse, colour, 4)
            label_text = f'{ellipse_area / 1000:.0f}'
            # cv2.putText(image, label_text, (int(x - len(label_text) * 10), int(y)), cv2.FONT_HERSHEY_SIMPLEX, 1, colour, 2)
            ellipses.append([x, y, a, b, theta])
print(f'{len(ellipses)} ellipses found and labelled.')
ellipses = np.array(ellipses)

# How to find 'holes' in the grid, where not all ellipses are found by this algorithm?
# Assign some row and column numbers to each ellipse
avg_ab = np.mean(ellipses[:, 2:4].flat)  # average of a and b ellipse params
# bigger gap in X direction though
distance = avg_ab * np.array([1.25, 0.9])
round_xy = np.round(ellipses[:, 0:2] / distance).astype(int)
# round_xy -= np.min(round_xy, axis=0)
# remove duplicates
unique_xy, indices = np.unique(round_xy, axis=0, return_index=True)
ellipses = ellipses[indices]
# for y in np.unique(round_xy[:, 1]):
#     find_row = unique_xy[:, 1] == y
#     this_row = unique_xy[find_row]
#     row_xlist = ellipses[find_row, 0]
#     min_x, max_x = np.min(row_xlist), np.max(row_xlist)
#     print(min_x, max_x)

for xy, (x, y, a, b, theta) in zip(unique_xy, ellipses):
    cv2.ellipse(image, ((x, y), (a, b), theta), (0, 255, 255), 2)
    label_text = f'{xy}'
    cv2.putText(image, label_text, (int(x - len(label_text) * 10), int(y)), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2)

# find holes
x_max, y_max = np.max(round_xy, axis=0) + 1
x_min, y_min = np.min(round_xy, axis=0) + 1
grid = set((x, y) for y in range(y_min, y_max) for x in range(x_min, x_max - ((y - y_min) % 2)))
print(sorted(list(grid)))
missing = sorted(list(grid - set((x, y) for x, y in round_xy)))
print(f'{missing=}')
for x, y in missing:
    find_row = unique_xy[:, 1] == y
    this_row = unique_xy[find_row]
    row_xlist = ellipses[find_row, 0]
    poly_x = np.polyfit(this_row[:, 0], row_xlist, deg=1)
    x_pos = np.polyval(poly_x, x)
    poly_y = np.polyfit(this_row[:, 1], ellipses[find_row, 1], deg=4)
    y_pos = np.polyval(poly_y, y)
    # x_pos = distance[0] * x
    _, _, a, b, theta = np.average(ellipses[find_row], axis=0)
    cv2.ellipse(image, ((x_pos, y_pos), (a, b), theta), (255, 255, 0), 2)
    label_text = f'{x}, {y}'
    cv2.putText(image, label_text, (int(x_pos - len(label_text) * 10), int(y_pos)), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 0), 2)

cv2.imwrite('1_image.jpg', image)

exit()

a, b = 95, 105
x0, y0 = 1357, 946  # start at bottom right
dx, dy = 1156/8, 854/9  # estimated from image 1
theta = 75
params = np.zeros((85, 5))
i = 0
# final_image = np.zeros((height, width, 3), np.uint8)
for r in range(10):
    y = int(y0 - r * dy)
    even = r / 2 == int(r / 2)
    x = x0 - (0 if even else 0.5 * dx)

    for c in range(9 if even else 8):
        print(x, y)
        # any matching from ellipse-fitting algo above?
        for (xf, yf), (af, bf), thetaf in ellipses:
            if (x - xf)**2 + (y - yf)**2 < a * b / 9:
                # x, y = xf, yf
                x, y, a, b, theta = xf, yf, af, bf, thetaf
                ellipse_area = pi * a * b
                print(f'Found at ({x:.0f}, {y:.0f}); (a, b) = ({a:.0f}, {b:.0f}), {theta=:.0f}, {ellipse_area=:.0f}')
                break

        def overlap(args):
            xx, yy, aa, bb, th = args
            build_image = np.zeros((height, width, 3), np.uint8)
            ellipse_params = ((xx, yy), (a, b), th)
            # print(ellipse_params)
            cv2.ellipse(build_image, ellipse_params, (255, 255, 255), 6)
            and_image = build_image[:, :, 0] & thresh
            cv2.imwrite('1_build.jpg', build_image)
            cv2.imwrite('1_and.jpg', and_image)
            result = sum(and_image.flat)
            # print([int(arg) for arg in args], result)
            return float(-result)

        start_vals = [x, y, a, b, theta]
        # print(overlap(x0))
        # x0[0] += 1.5
        # print(x0, overlap(x0))
        # res = scipy.optimize.basinhopping(overlap, x0, T=1e4, stepsize=10, niter=3,)
        # minimizer_kwargs={'method': 'BFGS', 'options': {'eps': 3}})
        res = scipy.optimize.minimize(overlap, start_vals, method='Nelder-Mead', tol=10)  # , options={'eps': 3}
        # res = scipy.optimize.least_squares(overlap, start_vals, diff_step=0.1)
        params[i] = res.x
        x, y, a, b, theta = res.x  # use a, b, theta as start vals for next ellipse
        ellipse_params = ((x, y), (a, b), theta)
        ellipse_area = pi * a * b
        print(f'Revised to ({x:.0f}, {y:.0f}); (a, b) = ({a:.0f}, {b:.0f}), {theta=:.0f}, {ellipse_area=:.0f}')
        cv2.ellipse(image, ellipse_params, (255, 0, 0), 4)
        x -= dx
        i += 1
        cv2.imwrite('1_image.jpg', image)

print(ellipses)