#include "widget.h"
#include "ui_widget.h"

#include <windows.h>
#include <QProcess>
#include <QWindow>
#include <QFileDialog>
#include <QtConcurrent>
#include <QDateTime>

Widget::Widget(QWidget *parent)
    : QWidget(parent)
    , ui(new Ui::Widget)
{
    ui->setupUi(this);
    // ui->widget
    ui->axWidget->dynamicCall("setCaption(string)", "ocx test"); //设置标题
    ui->axWidget->dynamicCall("setMenubar(bool)", false); //隐藏菜单栏
    ui->axWidget->dynamicCall("createNew(string)", "Word.Document"); //新建word空白文档
    ui->axWidget->setProperty("Visible", false); //新建word空白文档

    connect(this, SIGNAL(onTest()), this, SLOT(onTestDo()));
}
//    QProcess *pProc = new QProcess(this);
//    pProc->start("\"C:\Program Files (x86)\Microsoft Office\Office14\WINWORD.EXE");
//    pProc->waitForFinished(1000);
//    WId wid = Pid2Wid(pProc->processId(), "Microsoft Word");//获取窗口标识
//    QWindow *pWin = QWindow::fromWinId(wid);
//    QWidget *pWid = QWidget::createWindowContainer(pWin, this);
//    ui->verticalLayout->addWidget(pWid);


//    QFileDialog dialog;
//    dialog.setFileMode(QFileDialog::ExistingFile);
//    dialog.setViewMode(QFileDialog::Detail);
//    dialog.setOption(QFileDialog::ReadOnly, true);
//    dialog.setWindowTitle(QString("QAXwidget操作文件"));
//    dialog.setDirectory(QString("./"));
//    dialog.setNameFilter(QString("所有文件(*.*);;excel(*.xlsx);;word(*.docx *.doc);;pdf(*.pdf)"));
//    if(dialog.exec())
//    {
//        //根据文件后缀打开
//        QStringList files = dialog.selectedFiles();
//        for(auto filename : files)
//        {

//            if(filename.endsWith(".docx") || filename.endsWith(".doc"))
//            {
//                this->OpenWord(filename);
//            }
//        }
//    }
//}

void Widget::OpenWord(QString &filename)
{
    //  this->CloseOffice();
//    officeContent_ = new QAxWidget("Word.Application", this->ui->widget);
//    officeContent_->dynamicCall("SetVisible (bool Visible)", "false"); //不显示窗体
//    officeContent_->setProperty("DisplayAlerts", false);
//    //auto rect = this->ui->widget->geometry();

//    //officeContent_-> setGeometry(rect);
//    officeContent_->setControl(filename);
//    this->ui->verticalLayout->addWidget(officeContent_);
//    officeContent_->show();
//    m_word = officeContent_->querySubObject("ActiveDocument");
//    if(m_word)
//    {
//        m_app =   officeContent_->querySubObject("ActiveWindow");
//        if(m_app != NULL)
//        {


//        }

//    }

}


WId Widget::Pid2Wid(quint64 procID, const char *lpszClassName)
{

    char szBuf[256];
    HWND hWnd = GetTopWindow(GetDesktopWindow());
    while(hWnd)
    {
        DWORD wndProcID = 0;
        GetWindowThreadProcessId(hWnd, &wndProcID);

        if(wndProcID == procID)

        {
            GetClassNameA(hWnd, szBuf, sizeof(szBuf));

            if(strcmp(szBuf, lpszClassName) == 0)

            {

                return (WId)hWnd;

            }

        }

        hWnd = GetNextWindow(hWnd, GW_HWNDNEXT);

    }

    return 0;

}

void Widget::onTestDo()
{


    static bool testchangedValue = false;
    if(testchangedValue)
    {
        if(m_range)
        {
            // m_range->dynamicCall("Select()");
            m_range->setProperty("Text", QDateTime::currentDateTime().toString("yyyy-MM-dd hhmmsszzz"));
        }
    }
    else
    {
        if(m_range)
        {
            //m_range->dynamicCall("Select()");
            m_range->setProperty("Text", "1234");
        }
    }
    testchangedValue = !testchangedValue;

}


Widget::~Widget()
{
    while(1)
    {
        if(start == 0)
        {
            break;
        }
        if(start == 1)
        {
            start = 2;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(10));

    }
    delete ui;
}


void Widget::on_pushButton_clicked()
{
    m_word = ui->axWidget->querySubObject("ActiveDocument");
    if(m_word)
    {
        m_app =  m_word->querySubObject("ActiveWindow");
        if(m_app)
        {
            QAxObject *m_selection = m_app->querySubObject("Selection");
            if(m_selection)
            {
                m_range = m_selection->querySubObject("Range");
                if(m_range)
                {
                    m_range->setProperty("Text", "123");
                    if(start > 0)
                    {
                        return;
                    }
                    start = 1;
                    QtConcurrent::run([this]()
                    {
                        while(start == 1)
                        {
                            static long data = 0;
                            data++;
                            std::this_thread::sleep_for(std::chrono::milliseconds(300));
                            emit this->onTest();
                        }
                        start = 0;
                    });
                }
            }
        }
    }

}

void Widget::on_pushButton_2_clicked()
{
    m_word = ui->axWidget->querySubObject("ActiveDocument");
    if(m_word)
    {
        QAxObject   *m_rangeTmp = m_word->querySubObject("Content");
        auto int_start = m_rangeTmp->property("Start").toInt();
        m_range  = m_word->querySubObject("Range(QVariant, QVariant)", int_start, int_start + 1);
        if(m_range)
        {
            m_range->setProperty("Text", "123");
            if(start > 0)
            {
                return;
            }
            start = 1;
            QtConcurrent::run([this]()
            {
                while(start == 1)
                {
                    static long data = 0;
                    data++;
                    std::this_thread::sleep_for(std::chrono::milliseconds(300));
                    emit this->onTest();
                }
                start = 0;
            });
        }
    }
}
