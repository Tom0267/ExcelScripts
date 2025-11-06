import io
import sys
from PIL import Image
import plotly.io as pio

def convert_to_ico(svg_path, ico_path):
    with open(svg_path, 'r') as f:
        svg_content = f.read()

    # Convert SVG to PNG bytes using Kaleido
    png_bytes = pio.to_image({'data': [], 'layout': {'images': [{'source': svg_content}]}}, format='png')

    # Load into Pillow and save as ICO
    img = Image.open(io.BytesIO(png_bytes))
    img.save(ico_path, format='ICO', sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python iconify.py <FILENAME>")
    else:
        FILENAME = sys.argv[1]
        convert_to_ico(FILENAME + ".svg", FILENAME + ".ico")