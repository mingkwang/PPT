import os
import PIL.Image as Image
import win32com.client
rate = 1
image_infile = ''


def resize_by_width(infile, image_size):
    """按照宽度进行所需比例缩放"""
    im = Image.open(infile)
    (x, y) = im.size
    lv = round(x / image_size, 2) + 0.01
    x_s = int(x // lv)
    y_s = int(y // lv)
    out = im.resize((x_s, y_s), Image.ANTIALIAS)
    return out


def get_rate(infile):
    global rate
    im = Image.open(infile)
    (x, y) = im.size
    rate = y/x


def get_new_img_xy(infile, image_size):
    """返回一个图片的宽、高像素"""
    im = Image.open(infile)
    (x, y) = im.size
    lv = round(x / image_size, 2) + 0.01
    x_s = x // lv
    y_s = y // lv
    return x_s, y_s


# 定义图像拼接函数
def image_compose(image_colnum, image_size, image_rownum, image_names, image_save_path, x_new, y_new):
    to_image = Image.new('RGB', (image_colnum * x_new, image_rownum * y_new))  # 创建一个新图
    # 循环遍历，把每张图片按顺序粘贴到对应位置上
    total_num = 0
    for y in range(1, image_rownum + 1):
        for x in range(1, image_colnum + 1):
            from_image = resize_by_width(image_names[image_colnum * (y - 1) + x - 1], image_size)
            to_image.paste(from_image, ((x - 1) * x_new, (y - 1) * y_new))
            total_num += 1
            if total_num == len(image_names):
                break
    to_image.save(image_save_path[:-4]+'_min.jpg')  # 保存新图
    # print(image_save_path)
    get_big_images(image_infile, image_save_path, image_size, image_colnum, to_image)


def get_image_list_fullpath(dir_path):
    file_name_list1 = os.listdir(dir_path)
    file_name_list = []
    for i in range(1, len(file_name_list1)+1):
        if '幻灯片'+str(i)+'.jpg' in file_name_list1:
            file_name_list.append('幻灯片'+str(i)+'.jpg')
        if i >= 20:
            break
    image_fullpath_list = []
    for file_name_one in file_name_list:
        file_one_path = os.path.join(dir_path, file_name_one)
        if os.path.isfile(file_one_path):
            image_fullpath_list.append(file_one_path)
        else:
            img_path_list = get_image_list_fullpath(file_one_path)
            image_fullpath_list.extend(img_path_list)
    return image_fullpath_list


def get_big_images(image_file, image_save_path, image_size, image_colnum, to_image):
    # 添加第一页的大图
    # image1 = resize_by_width(file1, image_size*image_colnum+4)
    (x, y) = to_image.size
    image = Image.new('RGB', (image_size*image_colnum-4, int(y+image_size*image_colnum*rate)))
    from_image = resize_by_width(image_file, image_size*image_colnum)
    image.paste(from_image, (0, 0))
    image.paste(to_image, (0, int(image_size*image_colnum*rate)))
    image.save(image_save_path)
    # del_path = image_save_path[:-4]
    # if os.path.exists(del_path):
    #     os.remove(del_path)


def merge_images(image_dir_path, image_size, image_colnum):
    # 获取图片集地址下的所有图片名称
    image_fullpath_list = get_image_list_fullpath(image_dir_path)

    image_save_path = r'{}.jpg'.format(image_dir_path)  # 图片转换后的地址
    image_rownum_yu = len(image_fullpath_list) % image_colnum
    if image_rownum_yu == 0:
        image_rownum = len(image_fullpath_list) // image_colnum
    else:
        image_rownum = len(image_fullpath_list) // image_colnum + 1

    x_list = []
    y_list = []
    for img_file in image_fullpath_list:
        img_x, img_y = get_new_img_xy(img_file, image_size)
        x_list.append(img_x)
        y_list.append(img_y)

    x_new = int(x_list[len(x_list) // 5 * 4])
    y_new = int(x_list[len(y_list) // 5 * 4])

    image_compose(image_colnum, image_size, image_rownum, image_fullpath_list, image_save_path, x_new, int(x_new*rate))  # 调用函数


def ppt2png(ppt_path, filename, main_path):
    """
    ppt 转 png 方法
    :param ppt_path: ppt 文件的绝对路径
    :param long_sign: 是否需要转为生成长图的标识
    :return:
    """
    if os.path.exists(ppt_path):
        output_path = output_file(ppt_path, filename, main_path)  # 判断文件是否存在
        ppt_app = win32com.client.Dispatch('PowerPoint.Application')
        ppt = ppt_app.Presentations.Open(ppt_path)  # 打开 ppt
        ppt.SaveAs(output_path, 17)  # 17数字是转为 ppt 转为图片
        ppt_app.Quit()  # 关闭资源，退出

    else:
        raise Exception('请检查文件是否存在！\n')
    photo_path = main_path+'\\'+filename+'\\'+'photo'
    del_photo(photo_path)


def output_file(ppt_path, filename, path):
    """ 输出图片路径 """
    path = path + '\\' + filename
    if not os.path.exists(path):
        os.makedirs(path)
    output_png_path = os.path.join(path, 'photo.jpg')  # png 图片输出路径
    return output_png_path
    # file_name = os.path.basename(ppt_path)  # 获取文件名字
    # if file_name.endswith(('ppt', 'pptx')):
    #     exec_path = os.path.abspath(os.path.dirname(__file__))  # 当前脚本路径
    #     name = file_name.split('.')[0]  # 去除后缀，获取名字
    #     image_dir_path = os.path.join(exec_path, name)  # 图片文件夹的绝对路径
    #     if not os.path.exists(image_dir_path):
    #         os.makedirs(image_dir_path)  # 创建以 ppt 命名的图片文件夹
    #     output_png_path = os.path.join(image_dir_path, '一页一张图.png')  # png 图片输出路径
    #     return output_png_path
    # else:
    #     raise Exception('请检查后缀是否为 ppt/pptx 后缀！\n')


def del_photo(image_dir_path):
    global image_infile
    # image_dir_path = r'C:\Users\barrot\Desktop\test\01'  # 图片集地址
    image_infile = image_dir_path+'\\'+'幻灯片1.jpg'
    get_rate(image_infile)
    image_size = 256  # 每张小图片的大小
    image_colnum = 4  # 合并成一张图后，一行有几个小图
    merge_images(image_dir_path, image_size, image_colnum)


def process_bar(percent, start_str='', total_length=0):
    """
    进度条
    :return:   {:0>4.1f}%
    """
    bar = ''.join(["\033[1;31;41m%s\033[0m" % '   '] * int(percent * total_length)) + ''
    bar = '\r' + start_str + bar.ljust(total_length) + ' {:.2f}%'.format(percent * 100)
    print(bar, end='', flush=True)


if __name__ == '__main__':
    path = r'C:\Users\barrot\Desktop\1111'
    file_name_list1 = os.listdir(path)
    path_list = []
    flag = 0
    for reg in file_name_list1:
        reg1 = reg.split('.')
        flag += 1
        if 'ppt' in reg1 or 'PPT' in reg1 or 'pptx' in reg1 or 'PPTX' in reg1:
            reg = path+'\\'+reg
            ppt2png(reg, reg1[0], path)
            path_list.append(reg)
            process_bar(flag/len(file_name_list1), '转换进度：', 10)
    process_bar(1, '转换进度：', 10)
    print()
    print('转换完成，请到 '+path+' 的对应目录查看转换的图片')
