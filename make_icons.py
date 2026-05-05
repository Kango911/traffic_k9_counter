import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QPixmap, QPainter, QColor, QFont


def create_icon_png():
    # Создаём приложение Qt (необходимо для работы QPixmap)
    app = QApplication(sys.argv)

    pixmap = QPixmap(512, 512)
    pixmap.fill(QColor(131, 99, 157))
    painter = QPainter(pixmap)
    painter.setPen(QColor(255, 255, 255))
    font = QFont("Arial", 200, QFont.Bold)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), 0, "K9")
    painter.end()
    pixmap.save("k9_icon.png")
    print("Создан k9_icon.png. Теперь конвертируйте его в .ico и .icns онлайн или через конвертеры.")

    sys.exit(0)


if __name__ == "__main__":
    create_icon_png()