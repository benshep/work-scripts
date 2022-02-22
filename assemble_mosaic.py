import requests  # to download images
import os
from PIL import Image # to join images
from thunderforest import api_key


def assemble_mosaic():
    """Build a map from tiles using ThunderForest's API and export it to JPG and PNG files."""
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
    centre_x, centre_y = 16143, 10622

    # level = 14
    # centre_x, centre_y = 8073, 5311

    # level = 13  # doesn't show bike shops
    # centre_x, centre_y = 4035, 2655

    print(f'Creating image with {n_across} x {n_down} tiles, total resolution {overall_width} x {overall_height}')
    canvas = Image.new('RGB', (overall_width, overall_height), 'white')
    for i in range(n_across):
        x = centre_x - n_across // 2 + i
        print('\n', i, end=' ')
        for j in range(n_down):
            y = centre_y - n_down // 2 + j
            url = f"https://c.tile.thunderforest.com/cycle/{level}/{x}/{y}@2x.png?apikey={api_key}"
            print(j, end=' ')
            im = Image.open(requests.get(url, stream=True).raw)
            canvas.paste(im, (i * tile_size, j * tile_size))

    filename = os.path.join(os.environ['UserProfile'], 'Downloads', f'big_cycle_map_{level}')
    canvas.save(f"{filename}.png")
    canvas.save(f"{filename}.jpg")


if __name__ == '__main__':
    assemble_mosaic()
