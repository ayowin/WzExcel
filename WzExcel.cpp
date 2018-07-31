
#include "WzExcel.h"

WzExcel::WzExcel(const QString &filename)
{
    qDebug()<<"构造函数";

    OleInitialize(0); //当前线程初始化COM库并设置并发模式STA（single-thread apartment——单线程单元）

    excel = NULL;
    workBooks = NULL;
    workBook = NULL;
    workSheets = NULL;
    workSheet = NULL;
    data = NULL;

    this->filename = filename;
}

WzExcel::~WzExcel()
{
    qDebug()<<"析构函数";

    delete data;
    data = NULL;
    delete workSheet;
    workSheet = NULL;
    delete workSheets;
    workSheets = NULL;

    if(workBook != NULL)
    {
        workBook->dynamicCall("Close()");
        delete workBook;
        workBook = NULL;
    }
    if(workBooks != NULL)
    {
        workBooks->dynamicCall("Close()");
        delete workBooks;
        workBooks = NULL;
    }
    if (excel != NULL)
    {
        excel->dynamicCall("Quit()");
        delete excel;
        excel = NULL;
    }

    OleUninitialize(); //关闭当前线程的COM库并释放相关资源
}

bool WzExcel::open()
{
    qDebug()<<"open()";

    if(filename.isEmpty())
    {
        qDebug()<<"打开失败，文件名为空";
        return false;
    }

    excel = new QAxObject("Excel.Application"); //初始化excel对象
    excel->dynamicCall("SetVisible(bool)", false); //false不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示警告。
    workBooks = excel->querySubObject("WorkBooks"); //获取全部工作簿对象

    QFile file(filename);
    if (file.exists())
    {
        //导入文件到全部工作簿对象中，并将其设置为当前工作簿
        workBook = workBooks->querySubObject("Open(const QString &)", filename);
    }
    else
    {
        //文件不存在则创建
        workBooks->dynamicCall("Add");
        workBook = excel->querySubObject("ActiveWorkBook");
    }

    workSheets = workBook->querySubObject("Sheets"); //获得所有工作表对象
    return true;
}

bool WzExcel::setCurrentWorkSheet(const QString &sheetname)
{
    qDebug()<<"setCurrentWorkSheet()";

    if(filename.isEmpty())
    {
        qDebug()<<"设置当前工作表失败，文件名为空";
        return false;
    }

    if(sheetname.isNull())
    {
        workSheet = workSheets->querySubObject("Item(int)",1); //获得第一张工作表对象
    }
    else
    {
        int count = workSheets->property("Count").toInt();
        for(int i=1;i<=count;i++)
        {
            if(workSheets->querySubObject("Item(int)",i)->property("Name").toString() == sheetname)
            {
                workSheet = workSheets->querySubObject("Item(int)",i); //找到则设为指定名字的表对象
                break;
            }
            if(i==count)
            {
                workSheet = workSheets->querySubObject("Item(int)",1); //如果找不到则设为第一张表
            }
        }
    }
    qDebug()<<workSheet->property("Name").toString();
    return true;
}

bool WzExcel::createWorkSheet(const QString &sheetname)
{
    qDebug()<<"createWorkSheet()";

    if(filename.isEmpty())
    {
        qDebug()<<"创建工作表失败，文件名为空";
        return false;
    }

    int count = workSheets->property("Count").toInt();
    for(int i=1;i<=count;i++)
    {
        if(workSheets->querySubObject("Item(int)",i)->property("Name").toString() == sheetname)
        {
            qDebug()<<"该表已存在";
            return false;
        }
    }
    QAxObject *lastSheet = workSheets->querySubObject("Item(int)", count);
    workSheets->querySubObject("Add(QVariant)", lastSheet->asVariant());
    QAxObject *newSheet = workSheets->querySubObject("Item(int)", count);
    lastSheet->dynamicCall("Move(QVariant)", newSheet->asVariant());
    newSheet->setProperty("Name", sheetname);
    return true;
}

bool WzExcel::deleteWorkSheet(const QString &sheetname)
{
    qDebug()<<"deleteWorkSheet()";

    if(filename.isEmpty())
    {
        qDebug()<<"删除工作表失败，文件名为空";
        return false;
    }

    int count = workSheets->property("Count").toInt();
    for(int i=1;i<=count;i++)
    {
        if(workSheets->querySubObject("Item(int)",i)->property("Name").toString() == sheetname)
        {
            workSheets->querySubObject("Item(int)",i)->dynamicCall("delete");
            break;
        }
    }
    return true;
}

QString WzExcel::getValue(const int &row, const int &column)
{
    qDebug()<<"getValue()";

    if(filename.isEmpty())
    {
        qDebug()<<"获得指定位置的单元格内容失败，文件名为空";
        return NULL;
    }

    data = workSheet->querySubObject("Cells(int,int)", row, column);
    return data->dynamicCall("Value()").toString();
}

bool WzExcel::insertValue(const int &row, const int &column, const QString &value)
{
    qDebug()<<"insertValue()";

    if(filename.isEmpty())
    {
        qDebug()<<"插入指定位置的单元格内容失败，文件名为空";
        return false;
    }

    data = workSheet->querySubObject("Cells(int,int)", row, column);
    data->dynamicCall("Value", value);
    return true;
}

bool WzExcel::save()
{
    qDebug()<<"save()";

    if(filename.isEmpty())
    {
        qDebug()<<"保存失败，文件名为空";
        return false;
    }

    QFile file(filename);
    if (file.exists())
    {
        //文件存在则保存
        workBook->dynamicCall("Save()");
    }
    else
    {
        //文件不存在则另存为
        this->saveAs(filename);
    }
    return true;
}

bool WzExcel::saveAs(const QString &filename)
{
    qDebug()<<"saveAs()";

    if(filename.isEmpty())
    {
        qDebug()<<"另存为失败，文件名为空";
        return false;
    }

    workBook->dynamicCall("SaveAs(const QString &)",
                          QDir::toNativeSeparators(filename));
    return true;
}
