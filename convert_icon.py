from PIL import Image
import os

# Open the PNG file
img = Image.open('resolve_logo.png')

# Convert to RGBA if not already
if img.mode != 'RGBA':
    img = img.convert('RGBA')

# Save as ICO
img.save('resolve_logo.ico', format='ICO')
print("Icon converted successfully!") 