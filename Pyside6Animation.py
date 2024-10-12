from PySide6.QtWidgets import QApplication, QWidget,QLabel
from PySide6.QtCore import QPropertyAnimation, QPointF, Property
from PySide6.QtGui import QColor,QPainter,QRadialGradient
import sys
import math



class ColorWidget(QWidget):
    def __init__(self):
        super().__init__()
        self._color = QColor(255, 0, 0)  # Initial color

    def getColor(self):
        return self._color

    def setColor(self, color):
        self._color = color
        self.update()  # Trigger a repaint

    color = Property(QColor, getColor, setColor)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setBrush(self._color)
        painter.drawRect(self.rect())

class AnimatedColorWidget(ColorWidget):
    def __init__(self):
        super().__init__()
        self.initAnimation()

    def initAnimation(self):
        self.animation = QPropertyAnimation(self, b"color")
        self.animation.setDuration(5000)
        self.animation.setStartValue(QColor(255, 0, 0,))  # Start with red'
        self.animation.setKeyValueAt(0.5, QColor(0, 255, 0, ))
        self.animation.setEndValue(QColor(255, 0, 0,))  # End with blue
        self.animation.setLoopCount(-1)  # Loop indefinitely
        self.animation.start()

class RotatingGradientWidget(QWidget):
    def __init__(self):
        super().__init__()
        self._angle = 0
        self.initUI()
        self.initAnimation()

    def initUI(self):
        self.setGeometry(100, 100, 400, 400)
        self.setWindowTitle('Rotating Radial Gradient')

    def initAnimation(self):
        self.animation = QPropertyAnimation(self, b"angle")
        self.animation.setDuration(2000)
        self.animation.setStartValue(0)
        self.animation.setEndValue(360)
        self.animation.setLoopCount(-1)
        self.animation.start()

    def getAngle(self):
        return self._angle

    def setAngle(self, angle):
        self._angle = angle
        self.update()

    angle = Property(int, getAngle, setAngle)

    def paintEvent(self, event):
        painter = QPainter(self)
        gradient = QRadialGradient(QPointF(200, 200), 200)
        gradient.setColorAt(0, QColor(255, 0, 0))
        gradient.setColorAt(1, QColor(0, 0, 255))
        angle_in_radians = math.radians(float(self._angle))
        gradient.setCenter(QPointF(200 + 100 * math.cos(angle_in_radians),
                                   200 + 100 * math.sin(angle_in_radians)))
        painter.setBrush(gradient)
        painter.drawRect(self.rect())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    widget = AnimatedColorWidget()
    widget.resize(400, 300)
    widget.show()
    sys.exit(app.exec())

