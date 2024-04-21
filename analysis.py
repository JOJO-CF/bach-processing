import matplotlib.pyplot as plt
import pandas as pd

#为了显示中文，需要更换字体
plt.rcParams['font.family'] = ['sans-serif']
plt.rcParams['font.sans-serif'] = ['SimHei']


path = '蒙特卡洛分析收敛性分析（0.4）.csv'
df = pd.read_csv(path, usecols=['模型', '安全系数'])


# 计算纵轴抽取对应样本的概率
def fp(x, y):
    count = 0
    proba = []
    sample = x
    for i in range(len(x)):
            if y[i] <= 1:
                count = count + 1
                proba.append(round(count / (i+1), 3))
            else:
                proba.append(round(count / (i+1), 3))
    return sample, proba


sample, proba = fp(df['模型'], df['安全系数'])
plt.plot(sample, proba, color='black', linewidth=1, label='失效概率')  # 绘制直线，并给它一个我喜欢的颜色和宽度
plt.plot([0.2722]*5000, "r--", label='收敛值')  #绘制渐近线
plt.xlabel('样本数')
plt.ylabel('失效概率')
plt.legend(loc='upper right')  # 图例位置
# plt.title('蒙特卡洛模拟的收敛')
plt.show() #显示图表

