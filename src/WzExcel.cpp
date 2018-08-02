
#include "WzExcel.h"

WzExcel::WzExcel()
{
    qDebug()<<"WzExcel: WzExcel()";

    OleInitialize(0); //当前线程初始化COM库并设置并发模式STA（single-thread apartment——单线程单元）

    excel = NULL;
    workBooks = NULL;
    workBook = NULL;
    workSheets = NULL;
    workSheet = NULL;
    data = NULL;
    isOpened = false;
    fileName = "";
}

WzExcel::WzExcel(const QString &fileName)
{
    qDebug()<<"WzExcel: WzExcel(const QString &filename)";

    OleInitialize(0); //当前线程初始化COM库并设置并发模式STA（single-thread apartment——单线程单元）

    excel = NULL;
    workBooks = NULL;
    workBook = NULL;
    workSheets = NULL;
    workSheet = NULL;
    data = NULL;
    isOpened = false;

    this->fileName = fileName;
}

WzExcel::~WzExcel()
{
    qDebug()<<"WzExcel: ~WzExcel()";

    release();

    OleUninitialize(); //关闭当前线程的COM库并释放相关资源
}

void WzExcel::setFileName(const QString fileName)
{
    qDebug()<<"WzExcel: setFileName(const QString fileName)";

    this->fileName = fileName;
}

bool WzExcel::open(bool visible,bool displayAlerts)
{
    qDebug()<<"WzExcel: open(bool visible,bool displayAlerts)";

    if(fileName.isEmpty())
    {
        qDebug()<<"打开失败，文件名为空，请设置文件名，";
        return false;
    }

    excel = new QAxObject("Excel.Application"); //初始化excel对象
    if(excel == NULL)
    {
        qDebug()<<"打开失败，创建excel对象失败";
        return false;
    }
    excel->dynamicCall("SetVisible(bool)", visible); //false不显示窗体
    excel->setProperty("DisplayAlerts", displayAlerts); //不显示警告。
    workBooks = excel->querySubObject("WorkBooks"); //获取全部工作簿对象

    QFile file(fileName);
    if (file.exists())
    {
        //导入文件到全部工作簿对象中，并将其设置为当前工作簿
        workBook = workBooks->querySubObject("Open(const QString &)", fileName);
    }
    else
    {
        //文件不存在则创建
        workBooks->dynamicCall("Add");
        workBook = excel->querySubObject("ActiveWorkBook");
    }

    workSheets = workBook->querySubObject("Sheets"); //获得所有工作表对象
    isOpened = true;
    return true;
}

void WzExcel::close()
{
    qDebug()<<"WzExcel: close()";

    release();
}

bool WzExcel::setVisible(bool visible)
{
    qDebug()<<"WzExcel: setVisible(bool visible)";

    if(!isOpened)
    {
        qDebug()<<"设置visible失败，文件没有打开，请先调用open函数";
        return false;
    }
    else
    {
        excel->dynamicCall("SetVisible(bool)", visible);
        return true;
    }
}

bool WzExcel::setCurrentWorkSheet(const QString &sheetName)
{
    qDebug()<<"WzExcel: setCurrentWorkSheet(const QString &sheetName)";

    if(!isOpened)
    {
        qDebug()<<"设置当前工作表失败，文件没有打开，请先调用open函数";
        return false;
    }

    if(sheetName.isNull())
    {
        workSheet = workSheets->querySubObject("Item(int)",1); //获得第一张工作表对象
    }
    else
    {
        int count = workSheets->property("Count").toInt();
        for(int i=1;i<=count;i++)
        {
            if(workSheets->querySubObject("Item(int)",i)->property("Name").toString() == sheetName)
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

bool WzExcel::createWorkSheet(const QString &sheetName)
{
    qDebug()<<"WzExcel: createWorkSheet(const QString &sheetName)";

    if(!isOpened)
    {
        qDebug()<<"创建工作表失败，文件没有打开，请先调用open函数";
        return false;
    }

    int count = workSheets->property("Count").toInt();
    for(int i=1;i<=count;i++)
    {
        if(workSheets->querySubObject("Item(int)",i)->property("Name").toString() == sheetName)
        {
            qDebug()<<"该表已存在";
            return false;
        }
    }
    QAxObject *lastSheet = workSheets->querySubObject("Item(int)", count);
    workSheets->querySubObject("Add(QVariant)", lastSheet->asVariant());
    QAxObject *newSheet = workSheets->querySubObject("Item(int)", count);
    lastSheet->dynamicCall("Move(QVariant)", newSheet->asVariant());
    newSheet->setProperty("Name", sheetName);
    return true;
}

bool WzExcel::deleteWorkSheet(const QString &sheetName)
{
    qDebug()<<"WzExcel: deleteWorkSheet(const QString &sheetName)";

    if(!isOpened)
    {
        qDebug()<<"删除工作表失败，文件没有打开，请先调用open函数";
        return false;
    }

    int count = workSheets->property("Count").toInt();
    for(int i=1;i<=count;i++)
    {
        if(workSheets->querySubObject("Item(int)",i)->property("Name").toString() == sheetName)
        {
            workSheets->querySubObject("Item(int)",i)->dynamicCall("delete");
            break;
        }
    }
    return true;
}

QString WzExcel::getValue(const int &row, const int &column)
{
    qDebug()<<"WzExcel: getValue(const int &row, const int &column)";

    if(!isOpened)
    {
        qDebug()<<"获得指定位置的单元格内容失败，文件没有打开，请先调用open函数";
        return "";
    }

    if(workSheet == NULL)
    {
        qDebug()<<"获得指定位置的单元格内容失败，workSheet对象为空，请先调用setCurrentWorkSheet函数";
        return "";
    }

    data = workSheet->querySubObject("Cells(int,int)", row, column);
    return data->dynamicCall("Value()").toString();
}

bool WzExcel::insertValue(const int &row, const int &column, const QString &value)
{
    qDebug()<<"WzExcel: insertValue(const int &row, const int &column, const QString &value)";

    if(!isOpened)
    {
        qDebug()<<"插入指定位置的单元格内容失败，文件没有打开，请先调用open函数";
        return false;
    }

    if(workSheet == NULL)
    {
        qDebug()<<"插入指定位置的单元格内容失败,workSheet对象为空，请先调用setCurrentWorkSheet函数";
        return false;
    }

    data = workSheet->querySubObject("Cells(int,int)", row, column);
    data->dynamicCall("Value", value);
    return true;
}

bool WzExcel::save()
{
    qDebug()<<"WzExcel: save()";

    if(!isOpened)
    {
        qDebug()<<"保存失败，文件没有打开，请先调用open函数";
        return false;
    }

    QFile file(fileName);
    if (file.exists())
    {
        //文件存在则保存
        workBook->dynamicCall("Save()");
    }
    else
    {
        //文件不存在则另存为
        this->saveAs(fileName);
    }
    return true;
}

bool WzExcel::saveAs(const QString &fileName)
{
    qDebug()<<"WzExcel: saveAs(const QString &fileName)";

    if(!isOpened)
    {
        qDebug()<<"另存为失败，文件没有打开，请先调用open函数";
        return false;
    }

    workBook->dynamicCall("SaveAs(const QString &)",
                          QDir::toNativeSeparators(fileName));
    return true;
}

void WzExcel::release()
{
    isOpened = false;
    if(data != NULL)
    {
        delete data;
        data = NULL;
    }
    if(workSheet != NULL)
    {
        delete workSheet;
        workSheet = NULL;
    }
    if(workSheets != NULL)
    {
        delete workSheets;
        workSheets = NULL;
    }

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
}
