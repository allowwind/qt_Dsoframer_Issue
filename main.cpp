#include "widget.h"

#include <QApplication>
#include <QProcess>
#include <QWindow>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);



    Widget w;
    w.show();
    return a.exec();
}
