from PIL import Image

im = Image.open('EightBottom.png')

im.crop((0, 0, 840, 390)).save('EightBottom.png', quality=95)