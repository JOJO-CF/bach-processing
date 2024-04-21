# coding=gbk
import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
import math


#SPH����
def after_painter():
    plt.rcParams["font.sans-serif"] = ['SimHei']
    dataFolder = 'D:\\SPH(��ͬ��)\\������·��fs=0.5010��\\matlab����\\'
    ballRadius = 0.1

    # ����λ������
    postdistance = np.loadtxt(dataFolder + 'number230.txt')
    maxDist = 10
    minDist = 0
    fig = plt.figure(figsize=(20, 10))
    ax = fig.add_subplot(111)

    # ����λ����ͼ
    col = 1 - (postdistance[:, 4] - minDist) / (maxDist - minDist)
    if (maxDist - minDist) < 1e-5:
        col = np.ones(len(postdistance), 1)
    # print(col)
    ind = []
    for i in range(len(col)):
        ind.append(math.ceil((1 - col[i]) * 256))
        if ind[i] < 0:
            ind[i] = 0
    jet = plt.get_cmap('jet')
    for i in range(len(postdistance)):
        x1 = postdistance[i][0]
        y1 = postdistance[i][1]
        xy1 = (x1, y1)
        # print(xy1)
        rect = plt.Rectangle(xy1, 2 * ballRadius, 2 * ballRadius, angle=0.1, color=jet(ind[i]))
        ax.add_patch(rect)

    # plt.axis([0, 31, 0, 17])
    # plt.show()

    # ����������,colorbar
    plt.axis([0, 31, 0, 17])
    plt.xlabel('���¿�ȣ�m��', fontsize=12)
    plt.ylabel('���¸߶ȣ�m��', fontsize=12)
    plt.text(-4.7, 17.5, 'λ�ƣ�m��', size=12)
    position = fig.add_axes([0.1243, 0.925, 0.776, 0.0538])
    norm = mpl.colors.Normalize(vmin=0, vmax=10, clip=True)
    bounds = [round(elem, 2) for elem in np.linspace(0, 10, 11)]
    cb = fig.colorbar(mpl.cm.ScalarMappable(norm=norm, cmap=jet), ax=ax, cax=position, orientation='horizontal',
                      ticks=bounds)
    plt.show()


#SPH����У��
def check_slope():
    plt.rcParams["font.sans-serif"] = ['SimHei']
    dataFolder = 'D:\\SPH(��ͬ��)\\������·��fs=0.5010��\\matlabУ��\\'
    r = 0.1  # �޸ģ����Ӱ뾶
    dthreshold = 0.1
    a = np.loadtxt(dataFolder + 'formatlab.txt')
    fig = plt.figure()
    ax = fig.add_subplot(111)

    for i in range(len(a)):
        x = a[i][0]
        y = a[i][1]
        xy = (x, y)
        if a[i][2] == 1:
            rect1 = plt.Rectangle(xy, 2 * r, 2 * r, angle=0.1, color='red')
            ax.add_patch(rect1)
        if a[i][2] == 2:
            rect2 = plt.Rectangle(xy, 2 * r, 2 * r, angle=0.1, color='blue')
            ax.add_patch(rect2)
        if a[i][2] == 3:
            rect3 = plt.Rectangle(xy, 2 * r, 2 * r, angle=0.1, color='green')
            ax.add_patch(rect3)
        if a[i][2] == 4:
            rect4 = plt.Rectangle(xy, 2 * r, 2 * r, angle=0.1, color='yellow')
            ax.add_patch(rect4)

    plt.axis([0, 30.5, 0, 16.3])  # �޸ģ������᳤��
    plt.xlabel('���¿��/m', fontsize=18)
    plt.ylabel('���¸߶�/m', fontsize=18)
    plt.show()


# after_painter()                   #SPH����
check_slope()                     #SPH����У��