#ifndef WIDGET_H
#define WIDGET_H

#include <QAxObject>
#include <QAxWidget>
#include <QWidget>

QT_BEGIN_NAMESPACE
namespace Ui
{
    class Widget;
}
QT_END_NAMESPACE

class Widget : public QWidget
{
    Q_OBJECT

public:
    Widget(QWidget *parent = nullptr);
    ~Widget();

    void OpenWord(QString &filename);
private slots:
    void on_pushButton_clicked();

    void on_pushButton_2_clicked();

private:
    Ui::Widget *ui;
    int start = 0;
    WId Pid2Wid(quint64 procID, const char *lpszClassName);
    QAxWidget *officeContent_;
    QAxObject    *m_word = NULL;
    QAxObject    *m_app = NULL;

    QAxObject *m_range = NULL;

public slots:
    void onTestDo();



signals:
    void onTest();
};
#endif // WIDGET_H

