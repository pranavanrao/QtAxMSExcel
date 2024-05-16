#include "mainwindow.h"

#include <QApplication>
#include <QAxWidget>
#include <QMessageBox>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    MainWindow w;
    QAxWidget* excelWidget = new QAxWidget(&w);

    if (!excelWidget->setControl("{00020820-0000-0000-C000-000000000046}")) {
        QMessageBox::critical(nullptr, "Error", "Could not initiate Excel ActiveX control");
        return -1;
    }

    excelWidget->setGeometry(0, 0, 800, 600);

    QString filePath = "C:\\Users\\prana\\OneDrive\\Desktop\\QT-QML-projects\\QtAxMSExel\\students.xlsx"; // Replace with the path to your Excel file
    excelWidget->dynamicCall("Open(const QString&)", filePath);


    w.show();
    return a.exec();
}
