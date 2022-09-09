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
value = hsv[:, :, 2]

thresh = cv2.threshold(blue, 170, 255, cv2.THRESH_TOZERO_INV)[1]  # + cv2.THRESH_OTSU)[1]
thresh = cv2.threshold(thresh, 130, 255, cv2.THRESH_BINARY)[1]  # + cv2.THRESH_OTSU)[1]
thresh = (blue >= 120) & (blue <= 180)
# thresh = (value >= 120) & (blue <= 200)
thresh = thresh.astype(np.uint8) * 255
# thresh = cv2.Canny(thresh, 100, 200)
kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
cv2.imwrite('1_thresh.jpg', thresh)

# exit()

# Dilate with elliptical shaped kernel
dilate = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)
dilate = cv2.morphologyEx(dilate, cv2.MORPH_ERODE, kernel, iterations=2)
cv2.imwrite('1_dilate.jpg', dilate)
# dilate = cv2.dilate(thresh, kernel, iterations=2)

# Find contours, filter using contour threshold area, draw ellipse
# contours = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
# contours = contours[0] if len(contours) == 2 else contours[1]
# ellipse_count = 0
# for i, c in enumerate(contours):
#     colour = (randint(0, 255), randint(0, 255), randint(0, 255))
#     cv2.drawContours(image, [c], -1, colour, 4)
#     # compute the center of the contour
#     # M = cv2.moments(c)
#     # cX = M["m10"] / M["m00"]
#     # cY = M["m01"] / M["m00"]
#     # label_text = f'{i}'
#     # cv2.putText(image, label_text, (int(cX - len(label_text) * 10), int(cY)), cv2.FONT_HERSHEY_SIMPLEX, 1, colour, 2)
#     area = cv2.contourArea(c)
#     if len(c) >= 5:  # and 12000 > area > 1000:
#         ellipse = cv2.fitEllipse(c)
#         (x, y), (a, b), theta = ellipse
#         ellipse_area = pi * a * b
#         if 22000 < ellipse_area < 40000 and a / b > 0.5:
#             print(f'x, y = ({x:.0f}, {y:.0f}); (a, b) = ({a:.0f}, {b:.0f}), {theta=:.0f}, {ellipse_area=:.0f}')
#             ellipse_count += 1
#             colour = (255, 255, 255)  # (36, 255, 12)
#             cv2.ellipse(image, ellipse, colour, 4)
#             label_text = f'{ellipse_area / 1000:.0f}'
#             cv2.putText(image, label_text, (int(x - len(label_text) * 10), int(y)), cv2.FONT_HERSHEY_SIMPLEX, 1, colour, 2)
# print(f'{ellipse_count} ellipses found and labelled.')

a0, b0 = 95, 105
x0, y0 = 209, height - 144
d = 155
th0 = 75
params = np.zeros((85, 5))
i = 0
for r in range(10):
    even = r / 2 == int(r / 2)
    for c in range(9 if even else 8):
        x = x0 + c * d + (0 if even else 0.5 * d)
        y = int(y0 - r * 0.6 * d)

        def overlap(args):
            xx, yy, a, b, th = args
            build_image = np.zeros((height, width, 3), np.uint8)
            ellipse_params = ((xx, yy), (a, b), th)
            # print(ellipse_params)
            cv2.ellipse(build_image, ellipse_params, (255, 255, 255), 6)
            and_image = build_image[:, :, 0] & thresh
            return float(-sum(and_image.flat))

        print(overlap(np.zeros(5)))
        print(overlap(np.ones(5)))
        x0 = [x, y, a0, b0, th0]
        print(x0)
        res = scipy.optimize.minimize(overlap, x0)  # , method='Nelder-Mead')
        print(res)
        i += 1
        exit()
# cv2.imwrite('1_build.jpg', build_image)
# cv2.imwrite('1_and.jpg', and_image)

cv2.imwrite('1_image.jpg', image)
