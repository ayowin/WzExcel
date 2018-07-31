
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
    //传入需要操作的Excel文件的【绝对路径】
    WzExcel(const QString &filename);
    ~WzExcel();

    //打开,不存在则创建,成功返回true，失败返回false
    bool open();

    //设置当前工作表,成功返回true，失败返回false
    bool setCurrentWorkSheet(const QString &sheetname=NULL);

    //创建工作表，成功返回true，失败返回false
    bool createWorkSheet(const QString &sheetname);

    //删除工作表，成功返回true，失败返回false
    bool deleteWorkSheet(const QString &sheetname);

    //获得指定位置的单元格内容,成功返回该值，失败返回NULL
    QString getValue(const int &row,const int &column);

    //插入指定位置的单元格内容，成功返回true，失败返回false
    bool insertValue(const int &row,const int &colum,const QString &value);

    //保存，成功返回true，失败返回false
    bool save();

    //另存为，成功返回true，失败返回false
    bool saveAs(const QString &filename);

private:
    QString filename; //文件名

    QAxObject *excel; //excel对象
    QAxObject *workBooks; //所有工作簿对象
    QAxObject *workBook; //当前工作簿对象
    QAxObject *workSheets; //所有工作表对象
    QAxObject *workSheet; //当前工作表对象
    QAxObject *data; //当前数据对象
};


#endif
