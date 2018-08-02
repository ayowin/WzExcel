
#ifndef _WZEXCEL_H
#define _WZEXCEL_H

#include <QAxObject>
#include <QString>
#include <QDir>
#include <QFile>
#include <QDebug>
#include <windows.h>

/*
 *  类：WzExcel
 *  作用：Qt操作Excel
 *  作者：欧阳伟
 *  日期：2017-12-19
 *  用法示例(需在*.pro文件中添加：QT += axcontainer)
 *      WzExcel w("D:/hello.xlsx");                            //创建对象
 *      if(w.open())                                           //打开
 *      {
 *          w.setCurrentWorkSheet();                           //设置当前工作表
 *          for(int i=1;i<10;i++)
 *          {
 *              for(int j=1;j<10;j++)
 *              {
 *                  w.insertValue(i,j,QString::number(i*j));   //修改内容
 *              }
 *          }
 *          w.save();                                          //保存
 *          w.saveAs("D:/hello1.xlsx");                        //另存为
 *      }
 *
 *  说明：创建对象需传入绝对路径。
 */

class WzExcel
{
public:
    WzExcel();
    //传入需要操作的Excel文件的【绝对路径】
    WzExcel(const QString &fileName);
    ~WzExcel();

    //设置文件名
    void setFileName(const QString fileName);

    //打开,不存在则【新建一个工作簿(保存时以传入时的文件名保存)】,成功返回true，失败返回false
    bool open(bool visible=false,bool displayAlerts=false);

    //关闭
    void close();

    //设置visible，true: 可视，false: 隐藏
    bool setVisible(bool visible);

    //设置当前工作表,成功返回true，失败返回false
    bool setCurrentWorkSheet(const QString &sheetName=NULL);

    //创建工作表，成功返回true，失败返回false
    bool createWorkSheet(const QString &sheetName);

    //删除工作表，成功返回true，失败返回false
    bool deleteWorkSheet(const QString &sheetName);

    //获得指定位置的单元格内容,成功返回该值，失败返回NULL
    QString getValue(const int &row,const int &column);

    //插入指定位置的单元格内容，成功返回true，失败返回false
    bool insertValue(const int &row,const int &colum,const QString &value);

    //保存，成功返回true，失败返回false
    bool save();

    //另存为，成功返回true，失败返回false
    bool saveAs(const QString &fileName);

private:
    void release();

private:
    QString fileName; //文件名

    bool isOpened;

    QAxObject *excel; //excel对象
    QAxObject *workBooks; //所有工作簿对象
    QAxObject *workBook; //当前工作簿对象
    QAxObject *workSheets; //所有工作表对象
    QAxObject *workSheet; //当前工作表对象
    QAxObject *data; //当前数据对象
};


#endif
