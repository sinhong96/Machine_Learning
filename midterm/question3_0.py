import numpy as np
import matplotlib.pyplot as plt

# insert the input data (the data can be found in word file)
X = np.array([[1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
              [1, 0, 1, 1, 0, 1, 0, 0, 1, 1],
              [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
              [1, 1, 0, 1, 0, 0, 1, 1, 0, 1],
              [1, 0, 0, 1, 0, 0, 0, 0, 0, 1],
              [1, 0, 1, 0, 0, 0, 0, 0, 1, 0]])
# label
Y = np.array([1, -1, 1, -1, 1, -1])
# weight initialization，3 row 1 column(which means 3 inputs, 1 output)，take range from -1 to 1
# W = (np.zeros([10]) - 0.5) * 2
W = np.zeros([10])
print('The initialization Weight is:', W)
# learning rate
lr = 0.1
# epoch
n = 0
# Neural Network Output
O = 0


def update():
    global X, Y, W, lr, n
    n += 1
    O = np.dot(X, W.T)  # shap:(3,1)
    W_C = lr * ((Y - O.T).dot(X)) / int(X.shape[0])
    W = W + W_C


for i in range(3000):
    update()

# positive samples
x1 = [0, 1, 0]
y1 = [0, 1, 0]
z1 = [0, 1, 1]
# negative samples
x2 = [0, 1, 0]
y2 = [1, 0, 1]
z2 = [1, 1, 0]



print('The updated weight is:',W)
Z = np.dot(X, W.T)
print('The approximation result is:',Z)
