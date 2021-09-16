from PIL import Image # to join images
import requests  # to download images

# Central tile
#
# https://c.tile.thunderforest.com/cycle/15/16143/10621.png?apikey=a5dd6a2f1c934394bce6b0fb077203eb

aspect_ratio = 1 / 2**0.5  # aim for A0 paper
n_across = 28
n_down = round(n_across * aspect_ratio)
tile_size = 512
overall_width = tile_size * n_across
overall_height = tile_size * n_down

level = 15
centre_x = 16143
centre_y = 10622

# level = 14
# centre_x = 8073
# centre_y = 5311

# level = 13  # doesn't show bike shops
# centre_x = 4035
# centre_y = 2655

print(f'Creating image with {n_across} x {n_down} tiles, total resolution {overall_width} x {overall_height}')

canvas = Image.new('RGB', (overall_width, overall_height), 'white')
for ix, x in enumerate(range(int(centre_x - n_across / 2), int(centre_x + n_across / 2))):
    # if x != centre_x - 2: continue
    print('\n', ix, end=' ')
    for iy, y in enumerate(range(int(centre_y - n_down / 2), int(centre_y + n_down / 2))):
        url = f"https://c.tile.thunderforest.com/cycle/{level}/{x}/{y}@2x.png?apikey=a5dd6a2f1c934394bce6b0fb077203eb"
        print(iy, end=' ')
        im = Image.open(requests.get(url, stream=True).raw)
        canvas.paste(im, (ix * tile_size, iy * tile_size))

canvas.save(fr"C:\Users\bjs54\Downloads\big_cycle_map_{level}.png")
canvas.save(fr"C:\Users\bjs54\Downloads\big_cycle_map_{level}.jpg")
