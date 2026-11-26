"""Рисует иконку программы app.ico: фото товара в рамке и ценник (каталоги + прайсы)."""
import os
from PIL import Image, ImageDraw

S = 1024


def draw():
    im = Image.new('RGBA', (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    # фон — скруглённый квадрат, глубокий сине-зелёный
    d.rounded_rectangle((40, 40, S - 40, S - 40), radius=210, fill=(19, 78, 94))
    d.rounded_rectangle((40, 40, S - 40, S - 40), radius=210, outline=(39, 120, 138), width=18)
    # фотография (белая рамка, чуть повёрнута)
    photo = Image.new('RGBA', (560, 470), (0, 0, 0, 0))
    p = ImageDraw.Draw(photo)
    p.rounded_rectangle((0, 0, 559, 469), radius=36, fill=(255, 255, 255))
    p.rounded_rectangle((38, 38, 521, 380), radius=18, fill=(214, 236, 240))
    # силуэт кувшина/вазы на фото
    p.ellipse((190, 150, 370, 360), fill=(19, 78, 94))
    p.rectangle((235, 105, 325, 190), fill=(19, 78, 94))
    p.ellipse((222, 92, 338, 128), fill=(19, 78, 94))
    p.arc((310, 165, 415, 300), start=-80, end=80, fill=(19, 78, 94), width=26)
    p.ellipse((245, 190, 275, 260), fill=(90, 150, 165))   # блик
    photo = photo.rotate(-8, resample=Image.BICUBIC, expand=True)
    im.alpha_composite(photo, (150, 170))
    # ценник — тёплый оранжевый, с дыркой и шнурком
    tag = Image.new('RGBA', (400, 260), (0, 0, 0, 0))
    t = ImageDraw.Draw(tag)
    t.polygon([(0, 130), (90, 0), (399, 0), (399, 259), (90, 259)], fill=(245, 158, 11))
    t.ellipse((70, 108, 114, 152), fill=(19, 78, 94))
    # «строки цены»
    t.rounded_rectangle((160, 70, 350, 105), radius=16, fill=(255, 255, 255))
    t.rounded_rectangle((160, 150, 300, 185), radius=16, fill=(255, 245, 225))
    tag = tag.rotate(22, resample=Image.BICUBIC, expand=True)
    im.alpha_composite(tag, (470, 560))
    return im


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    im = draw()
    im.save(os.path.join(here, 'app.ico'), sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    im.resize((256, 256), Image.LANCZOS).save(os.path.join(here, 'app.png'))
    print('app.ico, app.png')


if __name__ == '__main__':
    main()
