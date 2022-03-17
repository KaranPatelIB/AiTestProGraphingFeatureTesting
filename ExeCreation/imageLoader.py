from PIL import Image
import os.path
import sys


def open_image(img_num):
    if 0 < img_num <= 3:

        dirName = os.path.dirname(__file__)
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        elif __file__:
            application_path = os.path.dirname(__file__)

        aiProLogo = os.path.join(application_path, 'images', 'aiProLogo.jpg')
        aiProTextLogo = os.path.join(application_path, 'images', 'aiProTextLogo.jpg')
        dgsmsLogo = os.path.join(application_path, 'images', 'dgsms.png')

        if img_num == 1:
            img = Image.open(aiProLogo)
        elif img_num == 2:
            img = Image.open(aiProTextLogo)
        elif img_num == 3:
            img = Image.open(dgsmsLogo)

        img.show()
        return True

    return False

