from PIL import Image
#
# image = Image.open('a.jpg')
# size = image.size
# img_ratio = size[0] / size[1]
# print(img_ratio)
# re_width = 600
# new_image = image.resize((re_width, int(re_width / img_ratio)))
# new_image.save('a.jpg')

def resize_image(img_path, new_width=100):
    image = Image.open(img_path)
    size = image.size
    img_ratio = size[0] / size[1]
    new_image = image.resize((new_width, int(new_width / img_ratio)))
    new_image.save(img_path)

resize_image('img\\a.jpg')