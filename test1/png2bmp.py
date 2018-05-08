import random
import os
import os.path as osp
from PIL import Image

def png2bmp(dataset_dir):
    def file_name(file_dir):
        L = []
        for root, dirs, files in os.walk(file_dir):
            for file in files:
                if os.path.splitext(file)[1] == '.png':
                    # L.append(os.path.join(root, file))
                    L.append(file)
        return L

    my_filename = file_name(dataset_dir)

    # random.shuffle(my_filename)
    # f = open('all.txt', 'w')
    # for i in my_filename:
    #     # k=' '.join([str(j) for j in i])
    #     f.write(i + "\n")
    # f.close()

    for i in my_filename:
        path = osp.join(dataset_dir, '%s' % i)
        img = Image.open(path)
        print("path=",path)
    img.save(path[0:len(path) - 4] + '.bmp')

if __name__ == '__main__':
    png2bmp('./')