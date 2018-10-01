#ifndef WORD_H
#define WORD_H

#include <QObject>


#include <QAxObject>
#include <QAxWidget>
#include <QAxBase>

#include <windows.h>
#include <QDir>


#include <QtDebug>

#include <QThread>

#include <QtCore/qmath.h>

#include <QAxWidget>

class MYWORD : public QObject
{
    Q_OBJECT
public:
   explicit MYWORD(QObject *parent = nullptr);

    ~MYWORD();

    ////
    QStringList listFiles;
    QList<QStringList> listPositionFindWord;
    QList<QAxObject*> listWordApplication;

    ///

    QStringList listMYWORD;

    QString FileDir,FileDir_S_R,FileDir_XP_XS_XW_X,FileDir_C_Z,FileDir_BQ,FileDir_DA,FileDir_DD,FileDir_FindMSWord;

    QString FileDir_U, FileDir_L, FileDir_TV;
    QString FileDir_HL_VD, FileDir_VT;

    QString saveDir;


    QAxObject* WordApplication, // Создаю интерфейс к MSWord  перечню

    *WordDocuments,  // Класс документа перечня
    *ActiveDocument, // Сделать документ активным
    *selection2,     // Создать класс Области страницы
    *Tables,         // Выбираем 1 таблицу в документе
    *StartCell,     // ячейка
    *CellRange;     // Область выбранной ячейки



    int temp;


    int columns; //Колонки в Перечне

    int rows; // Количество строчек в Перечне


    QStringList R; //резисторы
    QStringList RName; //имя резисторов


    QStringList C_Z;  //конденсаторы и фильтры
    QStringList C_ZName;  //имя конденсаторов
    QStringList C_NTD_PowerState; // По НТД постоянное напряжение только конденсаторов


    //КАРТА   РАБОЧИХ   РЕЖИМОВ   ЭЛЕКТРИЧЕСКИХ   СОЕДИНЕНИЙ,   ПРОВОДОВ   И   КАБЕЛЕЙ
    QStringList XP_XS_XW_X;  //Вилка
    QStringList XP_XS_XW_XName; //ИмяВилки

    //КАРТА   РАБОЧИХ   РЕЖИМОВ   КВАРЦЕВЫХ   РЕЗОНАТОРОВ,   КВАРЦЕВЫХ   МИКРОГЕНЕРАТОРОВ,   ПЬЕЗОЭЛЕКТРИЧЕСКИХ
    //И ЭЛЕКТРОМЕХАНИЧЕСКИХ   ФИЛЬТРОВ   И   ЛИНИЙ   ЗАДЕРЖКИ   НА   ПОВЕРХНОСТНЫХ   АКУСТИЧЕСКИХ   ВОЛНАХ
    QStringList BQ_G;
    QStringList BQ_GName;

    QStringList DD;
    QStringList DDName;

    QStringList DA;
    QStringList DAName;


    //РЕЖИМОВ   ВТОРИЧНЫХ   ИСТОЧНИКОВ   ПИТАНИЯ (ФОРМА  82/83)
    QStringList U;
    QStringList UName;


    QStringList L;
    QStringList LName;

    QStringList TV;
    QStringList TVName;

    QStringList  HL_VD;
    QStringList HL_VDName;

    QStringList  VT;
    QStringList VTName;



    //Берем элимент и название с данными
    QStringList Find_E;
    QStringList Find_EName;
    QList<QStringList> Find_Data_1;
    QList<QStringList> Find_Data_2;

    // То что нужно записать
    QList<QStringList> Send_Find_E;
    QStringList Send_Find_EName;
    QList<QStringList> Send_Find_Data_1;
    QList<QStringList> Send_Find_Data_2;

    QList<QStringList>  Send_Find_Data_1_1;
    QList<QStringList>  Send_Find_Data_2_2;

    QList<QStringList> list;
    QList<QStringList> list2;
    QStringList lol2;
    QStringList lol2_2;

    QStringList c_grm_codePower;
    QStringList c_grm_codeTemperatureRange;
    QStringList c_grm_power;
    QStringList c_grm_TemperatureRange;

    QStringList r_cr_code;
    QStringList r_cr_power;
    QStringList r_cr_TemperatureRange;
    QStringList r_cr_Void;

    QStringList z_nfm_code;
    QStringList z_nfm_power;

    QStringList c_avx_codePower;
    QStringList c_avx_power;
    QStringList c_avx_TemperatureRange;

    bool stateViewWord;

signals:
    void changeWork(int rows);

    void Part(QString);

    //QML
    void qml_StartFind();
    void qml_CreateShablon();
    void findData(int R ,int C_Z,int XP_XS_XW_X,int BQ_G,int DD,int DA,int U,int L,int TV,int HL_VD,int VT);

    void updateR(QString);
    void updateC_Z(QString);
    void updateXP_XS_XW_X(QString XP_XS_XW_X);
    void updateBQ_G(QString);
    void updateDD(QString);
    void updateDA(QString);
    void updateU(QString);
    void updateL(QString);
    void updateTV(QString);
    void updateHL_VD(QString);
    void updateVT(QString);

    ///////////
    void scanning(QString,int,int);

    void findWordFinish();

public slots:
    void process_start();

    void setTemp(QString);
    QString getTemp();

    void SetDirFindMSWord(QString);

    void OpenWord();

    void OpenWord_Perechen();

    void Findelements_Perechen();

    bool findRussianLanguage(QString text);


    void creatShablon();

    QString addData_R_Power_NTD(int i);
    QString addData_R_TemperatureRange_NTD(int i);
    QString addData_R_U_NTD(int i,double p);
    void CreatShablon_R();

    QString addData_C_Power_NTD(int i);
    QString addData_C_TemperatureRange_NTD(int i);
    void CreatShablon_C_Z();

    void CreatShablon_XP_XS_XW_X();

    void CreatShablon_VT();

    void CreatShablon_HL_VD();

    void CreatShablon_BQ();

    void CreatShablon_DA();

    void CreatShablon_DD();

    void CreatShablon_U();

    void CreatShablon_L();

    void CreatShablon_TV();

    void Work();

    void setViewFlag(int flag);
    int getViewFlag();


    //void WorkFind();



    //QML
    void qml_getFileName(QString str);
    QVariant qml_getlistMYWORD();
    QVariant qml_setChangeListMYWORD(QString index,QString shablonName);


    /////
    void scanDir(QDir dir);
    void scanDirWork(QString str,QDir dir);
    void findWord(QString str);
    void openWordFind(QString str,QString file);
    void openPDFFind(QString str,QString file);

    void closeAllWord();

};

#endif // WORD_H
