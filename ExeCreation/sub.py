from imageLoader import open_image

def user_img_inputs():
    img_num = int(input("\nSelect a image to open \n\nEnter a number(1, 2, or 3) : "))

    if open_image(img_num):
        return
    print("Invalid Selection")



