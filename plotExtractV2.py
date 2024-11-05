from pynput import mouse
from PIL import Image
from matplotlib import pyplot as plt


# Xmin = float(input('Input X-min: '))
# Xmax = float(input('Input X-max: '))
# Ymin = float(input('Input Y-min: '))
# Ymax = float(input('Input Y-max: '))
X = []
Y = []
NumberOfMouseClicks = 0
print('Click Origin')

def on_click(x, y, button, pressed):
    button = str(button)
    if button == "Button.right":
        listener.stop()
    #print(button, x, y)

    global NumberOfMouseClicks
    NumberOfMouseClicks = NumberOfMouseClicks + 1
    if NumberOfMouseClicks==1:
        print('Click Top Right')
    if NumberOfMouseClicks==3:
        print('Click data points. Right-click to end.')

    X.append(x)
    Y.append(y)

with mouse.Listener(on_click = on_click) as listener:
    listener.join()

X = X[::2]
Y = Y[::2]
del X[-1]
del Y[-1]
print(X, Y)