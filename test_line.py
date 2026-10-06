# -*- coding: utf-8 -*-
from PIL import Image
import numpy as np

img = Image.open(r'C:\Users\VKS\.gemini\antigravity\brain\53991c19-724d-4e5a-b4d4-c4a2d2eb6f14\.user_uploaded\media_1791087839970.png')
arr = np.array(img.convert('RGB'))

# Find green text Y
r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]
green_mask = (g > 65) & (g > r.astype(int) + 16) & (g > b.astype(int) + 16)
green_mask[:, int(img.width * 0.6):] = False
green_mask[:int(img.height * 0.15), :] = False
green_mask[int(img.height * 0.85):, :] = False
y_idx, _ = np.where(green_mask)
max_green_y = int(np.max(y_idx))

# Search for horizontal divider line below max_green_y
gray = np.mean(arr, axis=2)
divider_y = None

for y in range(max_green_y + 30, img.height - 10):
    row = gray[y, int(img.width * 0.1):int(img.width * 0.9)]
    # A horizontal line has gray values around 210-240, and is very uniform (low std dev)
    # Background is typically white (255)
    mean_val = np.mean(row)
    std_val = np.std(row)
    if mean_val < 245 and std_val < 8:
        divider_y = y
        break

print("Max green Y:", max_green_y)
print("Detected horizontal divider line Y:", divider_y)
print("Image total height:", img.height)
