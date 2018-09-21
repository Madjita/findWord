#include "word.h"


MYWORD::MYWORD(QObject *parent):
    QObject(parent),
    FileDir(""),
    temp(150),
    stateViewWord(true)
{

    saveDir = QDir::currentPath();

    connect(this, &MYWORD::qml_StartFind,this,&MYWORD::Work,Qt::QueuedConnection);
    connect(this, &MYWORD::qml_CreateShablon,this,&MYWORD::creatShablon,Qt::QueuedConnection);



    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/RShablon.docx");
    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/XPXSXWShablon.docx");
    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/CZShablon.docx");
    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/BQGShablon.docx");
    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/DAShablon.docx");
    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/DDShablon.docx");
    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/UShablon.docx");
    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/LShablon.docx");
    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/TVShablon.docx");
    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/VTShablon.docx");
    listMYWORD << QDir::currentPath()+QString("/Shablon")+QString("/HLVDShablon.docx");


    this->moveToThread(new QThread());
    connect(this->thread(),&QThread::started,this,&MYWORD::process_start);
    this->thread()->start();


}

MYWORD::~MYWORD()
{
    this->thread()->wait();
    delete this;
}

//сканировать папку
void MYWORD::scanDir(QDir dir)
{
    QStringList filters;
    filters << "*.docx"<< "*.doc" << "*.pdf";

    dir.setNameFilters(filters);
    dir.setFilter(QDir::Files | QDir::NoDotAndDotDot | QDir::NoSymLinks);

    // qDebug() << "Scanning: " << dir.path();
    QStringList fileList = dir.entryList();

    for (int i=0; i<fileList.count(); i++)
    {
        if(fileList[i] != "main.nut" &&
                fileList[i] != "info.nut")
        {
            //qDebug() << "Found file: " << fileList[i];

            listFiles <<  QString("%1/%2").arg(dir.absolutePath()).arg(fileList[i]);
        }

        // emit scanning(fileList[i],i,fileList.count());
    }

    dir.setFilter(QDir::AllDirs | QDir::NoDotAndDotDot | QDir::NoSymLinks);
    QStringList dirList = dir.entryList();
    for (int i=0; i<dirList.size(); ++i)
    {
        QString newPath = QString("%1/%2").arg(dir.absolutePath()).arg(dirList.at(i));
        scanDir(QDir(newPath));

        emit scanning(newPath,i,dirList.size());


    }


}

void MYWORD::scanDirWork(QString str,QDir dir)
{
    scanDir(dir);
    emit scanning("",0,1);

    findWord(str);
}



void MYWORD::findWord(QString str)
{
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);

    qDebug() << "СТАРТ";

    listPositionFindWord.clear();

    for (int i=0;i < listFiles.count();i++)
    {
          //if(listFiles[i].split('/').last().split('.').last() == "pdf")
              //openWordFind(str,listFiles[i]);
           // openPDFFind(str,listFiles[i]);

          openWordFind(str,listFiles[i]);
    }

     qDebug() << "Финиш";

}


void MYWORD::openWordFind(QString str, QString file)
{

    qDebug () << "DIR  = " << file;


    QFile File(file);
    File.setPermissions(File.permissions() | QFile::WriteOwner | QFile::WriteUser | QFile::WriteGroup | QFile::WriteOther);





    QAxObject* WordApplication = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    QAxObject* WordDocuments = WordApplication->querySubObject("Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    QAxObject *newDocument = WordDocuments->querySubObject("Add(QVariant)", QVariant(file));

   // WordDocuments->querySubObject( "Open(%T)",file);

    WordApplication->setProperty("Visible", 1);


    QAxObject* ActiveDocument = WordApplication->querySubObject("ActiveDocument()");
    //    QAxObject* content = ActiveDocument->querySubObject("Content");
    //    int mNumberOfPages = content->dynamicCall("Information(wdNumberOfPagesInDocument)").toInt();


    QAxObject* selection = WordApplication->querySubObject("Selection()");
    //selection->dynamicCall("HomeKey(wdStory)");

    QAxObject* selectionFind = selection->querySubObject("Find");
    //selectionFind->dynamicCall("ClearFormatting()");


    QList<QVariant> list2;
    list2.operator << (QVariant(str));
    list2.operator << (QVariant("0"));
    list2.operator << (QVariant("0"));
    list2.operator << (QVariant("0"));
    list2.operator << (QVariant("0"));
    list2.operator << (QVariant("0"));
    list2.operator << (QVariant(true));//true
    list2.operator << (QVariant("0"));
    list2.operator << (QVariant("0"));
    list2.operator << (QVariant("0"));//str + "12"
    list2.operator << (QVariant("0"));  //2
    list2.operator << (QVariant("0"));
    list2.operator << (QVariant("0"));
    list2.operator << (QVariant("0"));
    list2.operator << (QVariant("0"));


    //qDebug () << "mNumberOfPages = " << mNumberOfPages;

    qDebug () << "str = " << str;




    bool find = true;
    int countFindWord =0;

    QStringList listFindWord;

    while(find)
    {

        find = selectionFind->dynamicCall("Execute(const QVariant&,const QVariant&,"
                                          "const QVariant&,const QVariant&,"
                                          "const QVariant&,const QVariant&,"
                                          "const QVariant&,const QVariant&,"
                                          "const QVariant&,const QVariant&,"
                                          "const QVariant&,const QVariant&,"
                                          "const QVariant&,const QVariant&,const QVariant&)", list2).toBool();

        if(find)
        {
            //auto lol2 =  selectionFind->dynamicCall("Text");

            auto lol3 =  selection->querySubObject("Range()")->dynamicCall("Start");

            selection->querySubObject("Range()")->querySubObject("Font()")->querySubObject("Shading()")->setProperty("BackgroundPatternColor","wdColorBrightGreen");//ColorIndex

            countFindWord++;
            listFindWord.append(lol3.toString());
        }

    }

    qDebug() << countFindWord << "  ; " << listFindWord.count();

    if(countFindWord > 0)
    {
        listPositionFindWord.append(listFindWord);
    }
    else
    {
        WordApplication->dynamicCall("Quit (void)");
    }

    if(stateViewWord == false)
        WordApplication->dynamicCall("Quit (void)");

    delete WordApplication;


}

void MYWORD::openPDFFind(QString str, QString file)
{
    qDebug () << "DIR  = " << file;

    QAxObject * pdf = new QAxObject("Adobe PDF Reader");

    pdf->dynamicCall("Loadfile(%T)",file);

  //  QAxObject* pdf = new QAxObject("PDFCreator.clsPDFCreator"); // Создаю интерфейс к PDF


    //QAxObject* pdf = new QAxObject("Adobe PDF Reader"); // Создаю интерфейс к MSWord

  //  QAxObject* WordDocuments = WordApplication->querySubObject("Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

   // QAxObject *newDocument = WordDocuments->querySubObject("Add(QVariant)", QVariant(file));

   // pdf->querySubObject( "Loadfile(%T)",file);

   pdf->setProperty("Visible", 1);


   // QAxObject* ActiveDocument = WordApplication->querySubObject("ActiveDocument()");
    //    QAxObject* content = ActiveDocument->querySubObject("Content");
    //    int mNumberOfPages = content->dynamicCall("Information(wdNumberOfPagesInDocument)").toInt();



}



void MYWORD::process_start()
{
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
}

void MYWORD::setTemp(QString R)
{
    temp = R.toInt();
}

QString MYWORD::getTemp()
{
    return QString::number(temp);
}

void MYWORD::SetDirFindMSWord(QString data)
{
    FileDir_FindMSWord = data;
}


void MYWORD::Work()
{
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);

    OpenWord_Perechen();
}

void MYWORD::setViewFlag(int flag)
{
    switch (flag) {
    case 0: stateViewWord = false;break;
    case 2: stateViewWord = true;break;
    }

}

int MYWORD::getViewFlag()
{
    if(stateViewWord)
    {
        return 2;
    }

    return 0;
}

QVariant MYWORD::qml_getlistMYWORD()
{

    return  QVariant::fromValue(listMYWORD);
}

QVariant MYWORD::qml_setChangeListMYWORD(QString index, QString shablonName)
{
    QString shbalonNameFind = shablonName.split('/').last().split('.').first() ;
    QString shablonNameState = listMYWORD.value(index.toInt()).split('/').last().split('.').first();

    shablonName.remove(0,8);

    if(shbalonNameFind == shablonNameState)
    {
        listMYWORD.replace(index.toInt(),shablonName);
    }

    return  QVariant::fromValue(listMYWORD);
}

/*
void MYWORD::WorkFind()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    //  WordApplication_2->setProperty("Visible", 1);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject("Open(%T)",FileDir_FindMSWord); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");




    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");



    /////////////////////////////////////////////////////


    QAxObject* Tables_2,*StartCell_2,*CellRange_2,*StartCell_2_3,*CellRange_2_3;

    int flag =0;

    int k = 1;


    k = ActiveDocument_2->querySubObject("Tables")->property("Count").toInt();


    qDebug () << "K = " << k;



    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=1;i <= k;i++)
    {

        Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(i));


        // flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 2); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 2); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");



        QString text =  CellRange_2->property("Text").toString();


        Find_E.append(text);


        text = CellRange_2_3->property("Text").toString();

        Find_EName.append(text);

        Find_Data_1.append(QStringList());
        Find_Data_2.append(QStringList());

        //Cбор данных

       // int columns = Tables_2->querySubObject("Columns")->property("Count").toInt();

       // int rows = Tables_2->querySubObject("Rows")->property("Count").toInt();


        if(FileDir_FindMSWord.split('/').last() == "XPXSXW.docx")
        {


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 4); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 5); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 5); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 3); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 3); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 3); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 2); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 5); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 6); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 6); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 3); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////





            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 3); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");



            text =  CellRange_2->property("Text").toString();


            Find_E.append(text);


            text = CellRange_2_3->property("Text").toString();

            Find_EName.append(text);

            Find_Data_1.append(QStringList());
            Find_Data_2.append(QStringList());

            //Cбор данных

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 6); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 7); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 7); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 7); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 8); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 8); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 4); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 4); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");



            text =  CellRange_2->property("Text").toString();


            Find_E.append(text);


            text = CellRange_2_3->property("Text").toString();

            Find_EName.append(text);

            Find_Data_1.append(QStringList());
            Find_Data_2.append(QStringList());

            //Cбор данных

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 8); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 9); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 9); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 10); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 10); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 9); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 10); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 10); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 11); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 11); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////

            // Производим поиск собранных данных


            qDebug () << Find_E;
            qDebug () << Find_EName;

            qDebug () << Find_Data_1;

            qDebug () << Find_Data_2;

        }


        if(FileDir_FindMSWord.split('/').last() == "R.docx")
        {


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 4); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 4); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 4); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 3); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 5); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 5); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 5); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////





            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 3); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");



            text =  CellRange_2->property("Text").toString();


            Find_E.append(text);


            text = CellRange_2_3->property("Text").toString();

            Find_EName.append(text);

            Find_Data_1.append(QStringList());
            Find_Data_2.append(QStringList());

            //Cбор данных

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 6); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 6); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 6); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 7); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 7); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 7); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 4); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 4); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");



            text =  CellRange_2->property("Text").toString();


            Find_E.append(text);


            text = CellRange_2_3->property("Text").toString();

            Find_EName.append(text);

            Find_Data_1.append(QStringList());
            Find_Data_2.append(QStringList());

            //Cбор данных

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 8); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 8); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 8); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 9); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 9); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 9); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 10); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 10); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////

            // Производим поиск собранных данных


            qDebug () << Find_E;
            qDebug () << Find_EName;

            qDebug () << Find_Data_1;

            qDebug () << Find_Data_2;

        }



        qDebug ()  << "========================================================";

    }


    QStringList Send_Find_E_bekap,result;

    Send_Find_E_bekap = Find_EName;


    bool flagApp=true;

    bool flagApp2=true;


    int f = 0;
    int save_sov_i = 0;



    //    for(int i=0;i < Find_EName.count();i++)
    //    {
    //        if(result.count() < 1)
    //        {
    //            flagApp = true;

    //            result.append(Find_EName[i]);
    //            Send_Find_E.append(QStringList());
    //            Send_Find_E[result.count()-1].append(Find_E[i]);
    //            Send_Find_Data_1.append(Find_Data_1[i]);
    //            Send_Find_Data_2.append(Find_Data_2[i]);

    //        }
    //        else
    //        {

    //            int res = result.count();

    //            for(int j=0; j < res;j++)
    //            {
    //                flagApp = false;

    //                if(Find_EName[i] == result[j])
    //                {
    //                    flagApp = true;

    //                    break;
    //                }

    //            }

    //            if(flagApp == false)
    //            {
    //                result.append(Find_EName[i]);
    //                Send_Find_E.append(QStringList());
    //                Send_Find_E[result.count()-1].append(Find_E[i]);
    //                Send_Find_Data_1.append(Find_Data_1[i]);
    //                Send_Find_Data_2.append(Find_Data_2[i]);
    //            }
    //            else
    //            {
    //                Send_Find_E[result.count()-1].append(Find_E[i]);
    //                Send_Find_Data_1[result.count()-1].append(Find_Data_1[i]);
    //                Send_Find_Data_2[result.count()-1].append(Find_Data_2[i]);

    //            }

    //        }


    //    }

    QString first;
    do
    {

        flagApp = true;
        first = Find_EName[0];
        result.append(Find_EName[0]);
        Send_Find_E.append(QStringList());
        Send_Find_E[result.count()-1].append(Find_E[0]);
        Send_Find_Data_1.append(Find_Data_1[0]);
        Send_Find_Data_2.append(Find_Data_2[0]);

        for(int i=1;i < Find_EName.count();i++)
        {
            if(Find_EName[i] == first)
            {
                flagApp = false;
                Find_EName.removeAt(i);
                Send_Find_E[result.count()-1].append(Find_E[i]);
                Find_E.removeAt(i);
                Send_Find_Data_1[result.count()-1].append(Find_Data_1[i]);
                Find_Data_1.removeAt(i);
                Send_Find_Data_2[result.count()-1].append(Find_Data_2[i]);
                Find_Data_2.removeAt(i);
                i--;
            }
        }

        Find_EName.removeAt(0);
        Find_E.removeAt(0);
        Find_Data_1.removeAt(0);
        Find_Data_2.removeAt(0);

        first = "";




    }while(Find_EName.count() > 1);

    qDebug ()  << "========================================================";

    qDebug () << result;

    qDebug ()  << "======================111111===========================";

    qDebug () << Send_Find_Data_1;

    qDebug ()  << "====================2222222222============================";

    qDebug () << Send_Find_Data_2;

    qDebug ()  << "========================================================";

    qDebug () << Send_Find_E;

    qDebug ()  << "========================================================";



    bool flag_find_sovpad = false;

    QStringList result_2,Send_Find_Data_1_eshe,Send_Find_Data_2_eshe;
    QList<QStringList> Send_Find_E_eshe;

    for(int i=0 ;i < Send_Find_E.count();i++)
    {
        if(Send_Find_E[i].count() > 0)
        {

            for(int j=0; j < Send_Find_Data_1[i].count();j++)
            {

                lol2.append(Send_Find_Data_1[i].value(j));

                if(((j%11) == 0 ) && ((j > 0)&& (j <= 11)))
                {

                    list.append(lol2);
                    lol2.clear();
                }
                else
                {
                    qDebug() << QString::number(j%12);

                    if(((j%12) == 11 ) && (j > 11))
                    {
                        list.append(lol2);
                        lol2.clear();
                    }
                }

            }

            for(int j=0; j < Send_Find_Data_2[i].count();j++)
            {

                lol2_2.append(Send_Find_Data_2[i].value(j));

                if(((j%11) == 0 ) && ((j > 0)&& (j <= 11)))
                {

                    list2.append(lol2_2);
                    lol2_2.clear();
                }
                else
                {
                    qDebug() << QString::number(j%12);

                    if(((j%12) == 11 ) && (j > 11))
                    {
                        list2.append(lol2_2);
                        lol2_2.clear();
                    }
                }

            }

            //подумать !!!!
            for(int j=0;j < list.count()-1;j++)
            {
                if(  (list[j] != list[j+1]) ||  ( list2[j] != list2[j+1]))
                {
                    //  Send_Find_Data_1.replace(i,list[j]);
                    //  Send_Find_Data_2.replace(i,list2[j]);

                    //                    auto list_copy = list;

                    //                    QStringList first = list_copy[0];
                    //                    Send_Find_E.append(QStringList());

                    //                    do
                    //                    {
                    //                        for(int k =0;k < list_copy.count();k++)
                    //                        {
                    //                            if(list_copy[k] == first)
                    //                            {


                    //                                result.append(result[i]);

                    //                                Send_Find_E.last().append(Send_Find_E[i].value(j+1));
                    //                                Send_Find_E[i].removeAt(j+1);
                    //                                Send_Find_Data_1[result.count()-1].append(Find_Data_1[i]);
                    //                                Find_Data_1.removeAt(i);
                    //                                Send_Find_Data_2[result.count()-1].append(Find_Data_2[i]);
                    //                                Find_Data_2.removeAt(i);
                    //                                list_copy.removeAt(k);
                    //                                k--;
                    //                            }
                    //                        }

                    //                    }while(list_copy.count() < 1);

                    //  Send_Find_E.append(QStringList());
                    //  Send_Find_E.last().append(Send_Find_E[i].value(j+1));
                    //  Send_Find_E[i].removeAt(j+1);

                    Send_Find_E_eshe.append(QStringList());
                    Send_Find_E_eshe.last().append(Send_Find_E[i].value(j+1));
                    Send_Find_E[i].removeAt(j+1);

                    result_2.append(result[i]);

                    //  list.removeAt(j+1);
                    //   Send_Find_Data_1.append(list[j+1]);
                    //   Send_Find_Data_2.append(list2[j+1]);

                    list.removeAt(j+1);
                    Send_Find_Data_1_eshe.append(list[j+1]);
                    Send_Find_Data_2_eshe.append(list2[j+1]);

                    flag_find_sovpad = true;


                    break;
                }
                else
                {
                    //                    Send_Find_Data_1.replace(i,list[j]);
                    //                    Send_Find_Data_2.replace(i,list2[j]);
                }
            }


            list.clear();
            list2.clear();

            flag_find_sovpad = false;

        }
    }


    qDebug ()  << "========================================================";

    qDebug () << result;

    qDebug ()  << "======================111111===========================";

    qDebug () << Send_Find_Data_1;

    qDebug ()  << "====================2222222222============================";

    qDebug () << Send_Find_Data_2;

    qDebug ()  << "========================================================";

    qDebug () << Send_Find_E;

    qDebug ()  << "========================================================";


    //ActiveDocument_2->dynamicCall("Close (boolean)", false);

    WordApplication_2->dynamicCall("Quit (void)");


    /////////////////////////////////////END////////////////////////////////////////////////////////////////////////////


    if(FileDir_FindMSWord.split('/').last() == "XPXSXW.docx")
    {

        WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

        // WordApplication_2->setProperty("Visible", 1);

        WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

        WordDocuments_2->querySubObject( "Open(%T)",FileDir_XP_XS_XW_X); //D:\\11111\\One.docx


        ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



        // ActiveDocument_2->querySubObject("Range()")->dynamicCall("Copy()");


        ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");






        selection_2 = WordApplication_2->querySubObject("Selection()");


        qDebug() <<"Send_Find_E.count()/3 = " << QString::number(Send_Find_E.count()%3);



        if((((Send_Find_E.count()-1)%3) > 0)  && (((Send_Find_E.count()-1)%3) !=0))
        {
            for(int i=1; i < (Send_Find_E.count()/3)+1;i++)
            {

                selection_2->dynamicCall("EndKey(wdStory)");
                selection_2->dynamicCall("InsertBreak()");
                selection_2->dynamicCall("Paste()");

            }
        }
        else
        {
            for(int i=1; i < (Send_Find_E.count()/3);i++)
            {

                selection_2->dynamicCall("EndKey(wdStory)");
                selection_2->dynamicCall("InsertBreak()");
                selection_2->dynamicCall("Paste()");

            }
        }



        /////////////////////////////////////////////////////


        flag =0;


        k = 1;




        selection_2->dynamicCall("HomeKey(wdStory)");


        qDebug () << "K = " << k;

        Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


        selection_2->dynamicCall("HomeKey(wdStory)");


        QString text;

        for(int i =0 ; i < Send_Find_E.count();i++)
        {

            flag++;

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            if(Send_Find_E[i].count() < 1)
            {
                CellRange_2->dynamicCall("InsertAfter(Text)", Send_Find_E[i].value(0));
            }
            else
            {
                QString str = "";

                for(int j=0;j < Send_Find_E[i].count();j++)
                {
                    if(j != Send_Find_E[i].count()-1)
                    {
                        str +=Send_Find_E[i].value(j).split(0x000d).first()+", ";
                    }
                    else
                    {
                        str +=Send_Find_E[i].value(j);
                    }
                }
                CellRange_2->dynamicCall("InsertAfter(Text)", str);
            }

            CellRange_2_3->dynamicCall("InsertAfter(Text)", result[i]);

            switch (flag) {

            case 1:
            {

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 4); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 5); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(1))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 5); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 2); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 5); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                }



                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 6); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 6); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();
                if(text != Send_Find_Data_2[i].value(2))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(3))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(4))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(5))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(6))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(7))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(8))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(9))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(10))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                }

                //    /////////////////////////////////////////////////////


                //////////////////////////////////////////////////////////////////////////////////////

                break;
            }
            case 2:
            {

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 6); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 7); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 7); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(2))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(3))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(4))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(5))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(6))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(7))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(8))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(9))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(10))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(11))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                }


                //////////////////////////////////////////////////////////////////////////////////////



                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 7); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 8); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 8); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                }

                //    /////////////////////////////////////////////////////

                break;
            }
            case 3:
            {


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 8); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(0))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 9); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(1))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 9); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 10); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 10); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = StartCell_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                }


                //    //////////////////////////////////////////////////////////////////////////////////////


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 9); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(0))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 10); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(1))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 10); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 11); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 11); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                }

                //    /////////////////////////////////////////////////////
                break;
            }

            }


            if(flag == 3)
            {
                flag =0;

                k++;
                if(k > (Send_Find_E.count()/3))
                {

                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                    qDebug () << "K = " << k;
                }

            }
        }





        //Сохранить pdf
        //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//XPXSXW" ,"17");//fileName.split('.').first()

        ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", "D://11111//1//XPXSXWRelase");


        //ActiveDocument_2->dynamicCall("Close (boolean)", false);

        WordApplication_2->dynamicCall("Quit (void)");

    }


    if(FileDir_FindMSWord.split('/').last() == "R.docx")
    {

        WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

        // WordApplication_2->setProperty("Visible", 1);

        WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

        WordDocuments_2->querySubObject( "Open(%T)",FileDir_S_R); //D:\\11111\\One.docx


        ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



        // ActiveDocument_2->querySubObject("Range()")->dynamicCall("Copy()");


        ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");






        selection_2 = WordApplication_2->querySubObject("Selection()");


        qDebug() <<"Send_Find_E.count()/3 = " << QString::number(Send_Find_E.count()%3);



        if((((Send_Find_E.count()-1)%3) > 0)  && (((Send_Find_E.count()-1)%3) !=0))
        {
            for(int i=1; i < (Send_Find_E.count()/3)+1;i++)
            {

                selection_2->dynamicCall("EndKey(wdStory)");

                selection_2->dynamicCall("InsertBreak()");


                selection_2->dynamicCall("Paste()");

            }
        }
        else
        {
            for(int i=1; i < (Send_Find_E.count()/3);i++)
            {

                selection_2->dynamicCall("EndKey(wdStory)");

                selection_2->dynamicCall("InsertBreak()");

                selection_2->dynamicCall("Paste()");

            }
        }



        /////////////////////////////////////////////////////


        flag =0;


        k = 1;




        selection_2->dynamicCall("HomeKey(wdStory)");


        qDebug () << "K = " << k;

        Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


        selection_2->dynamicCall("HomeKey(wdStory)");


        QString text;

        for(int i =0 ; i < Send_Find_E.count();i++)
        {

            flag++;

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            if(Send_Find_E[i].count() < 1)
            {
                CellRange_2->dynamicCall("InsertAfter(Text)", Send_Find_E[i].value(0));
            }
            else
            {
                QString str = "";

                for(int j=0;j < Send_Find_E[i].count();j++)
                {
                    if(j != Send_Find_E[i].count()-1)
                    {
                        str +=Send_Find_E[i].value(j).split(0x000d).first()+", ";
                    }
                    else
                    {
                        str +=Send_Find_E[i].value(j);
                    }
                }
                CellRange_2->dynamicCall("InsertAfter(Text)", str);
            }

            CellRange_2_3->dynamicCall("InsertAfter(Text)", result[i]);

            switch (flag) {

            case 1:
            {

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 4); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 4); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(1))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 4); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 5); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                }



                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 5); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 5); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();
                if(text != Send_Find_Data_2[i].value(2))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(3))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(4))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(5))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(6))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(7))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(8))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(9))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(10))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                }

                //    /////////////////////////////////////////////////////


                //////////////////////////////////////////////////////////////////////////////////////

                break;
            }
            case 2:
            {

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 6); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 6); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 6); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(2))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(3))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(4))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(5))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(6))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(7))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(8))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(9))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(10))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(11))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                }


                //////////////////////////////////////////////////////////////////////////////////////



                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 7); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 7); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 7); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                }

                //    /////////////////////////////////////////////////////

                break;
            }
            case 3:
            {


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 8); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(0))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 8); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(1))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 8); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = StartCell_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                }


                //    //////////////////////////////////////////////////////////////////////////////////////


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 9); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(0))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 9); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(1))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 9); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 10); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 10); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                }

                //    /////////////////////////////////////////////////////
                break;
            }

            }


            if(flag == 3)
            {
                flag =0;

                k++;
                if(k > (Send_Find_E.count()/3))
                {

                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                    qDebug () << "K = " << k;
                }

            }
        }





        //Сохранить pdf
        //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//XPXSXW" ,"17");//fileName.split('.').first()

        ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", "D://11111//1//RRelase");


        //ActiveDocument_2->dynamicCall("Close (boolean)", false);

        WordApplication_2->dynamicCall("Quit (void)");

    }



    result.clear();

    Send_Find_Data_1.clear();

    Send_Find_Data_2.clear();

    Find_E.clear();
    Find_EName.clear();
    Find_Data_1.clear();
    Find_Data_2.clear();

    // То что нужно записать
    Send_Find_E.clear();
    Send_Find_EName.clear();
    Send_Find_Data_1.clear();
    Send_Find_Data_2.clear();

    Send_Find_Data_1_1.clear();
    Send_Find_Data_2_2.clear();

    list.clear();
    lol2.clear();





}
*/
void MYWORD::qml_getFileName(QString str)
{
    FileDir = str;
}



void MYWORD::OpenWord()
{
    QAxObject* WordApplication = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    // WordApplication->setProperty("Visible", 1);

    QAxObject* WordDocuments = WordApplication->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments->querySubObject( "Open(%T)",FileDir); //D:\\11111\\One.docx


    QAxObject* ActiveDocument = WordApplication->querySubObject("ActiveDocument()");



    QAxObject *selection2 = WordApplication->querySubObject("Selection()");


    QAxObject* Tables = selection2->querySubObject("Tables(1)");



    QAxObject* StartCell  = Tables->querySubObject("Cell(Row, Column)", 6, 2); // (ячейка C1)
    QAxObject* CellRange = StartCell->querySubObject("Range()");



    //CellRange->dynamicCall("InsertAfter(Text)", "НУ");


    //    StartCell = Tables->querySubObject("Cell(Row, Column)", 8, 3);

    //    CellRange = StartCell->querySubObject("Range()");



    //    auto lol =  CellRange->property("Text");

    //    qDebug () << lol.toString();

    auto columns = Tables->querySubObject("Columns")->property("Count").toInt();

    auto rows = Tables->querySubObject("Rows")->property("Count").toInt();

    qDebug () << "Колонки = " << columns;

    qDebug () <<"Строки = " << rows;


    //////////////////////////////////////////////////////////////////////////////
    int count_find = 0;

    for(int i=1; i <  rows;i++)
    {
        for(int j=1; j < columns; j++)
        {

            StartCell = Tables->querySubObject("Cell(Row, Column)", i, j);

            CellRange = StartCell->querySubObject("Range()");

            QString text =  CellRange->property("Text").toString();

            if((text[0] == "R") && (j == 2))
            {
                count_find++;

            }
        }
    }

    qDebug () << QString::number(count_find);

    ///////////////////////////////////////////////////////////////////////////////


    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    // WordApplication_2->setProperty("Visible", 1);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",FileDir_S_R); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



    // ActiveDocument_2->querySubObject("Range()")->dynamicCall("Copy()");


    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");







    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");




    qDebug() <<"count_find/3 = " << QString::number(count_find/3);



    for(int i=0; i < (count_find/3);i++)
    {

        selection_2->dynamicCall("EndKey(wdStory)");
        selection_2->dynamicCall("InsertBreak()");
        selection_2->dynamicCall("Paste()");
    }



    //    QAxObject* Tables_2 = ActiveDocument_2->querySubObject("Tables(1)");



    //    QAxObject* StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 2); // (ячейка C1)
    //    QAxObject* CellRange_2 = StartCell_2->querySubObject("Range()");

    //    QAxObject* StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 2); // (ячейка C1)
    //    QAxObject* CellRange_2_3 = StartCell_2_3->querySubObject("Range()");




    /////////////////////////////////////////////////////


    QAxObject* Tables_2,*StartCell_2,*CellRange_2,*StartCell_2_3,*CellRange_2_3;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=1; i <  rows;i++)
    {
        for(int j=1; j < columns; j++)
        {

            StartCell = Tables->querySubObject("Cell(Row, Column)", i, j);

            CellRange = StartCell->querySubObject("Range()");

            QString text =  CellRange->property("Text").toString();

            if((text[0] == "R") && (j == 2))
            {

                flag++;

                StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
                CellRange_2 = StartCell_2->querySubObject("Range()");

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                CellRange_2->dynamicCall("InsertAfter(Text)", text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();

                CellRange_2_3->dynamicCall("InsertAfter(Text)", text);



                if(flag == 3)
                {
                    flag =0;

                    k++;
                    if(k > (count_find/3)+1 )
                    {

                        qDebug () << "Конец ; K = " << k;
                        break;
                    }
                    else
                    {
                        Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                        qDebug () << "K = " << k;
                    }

                }

            }
        }
    }






    //Сохранить pdf
    //  ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//Good" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT");//"D://11111//1//Good");
    //ActiveDocument_2->dynamicCall("Close (boolean)", false);
    ActiveDocument->dynamicCall("Close (boolean)", false);



    WordApplication->dynamicCall("Quit (void)");

    WordApplication_2->dynamicCall("Quit (void)");
}

void MYWORD::OpenWord_Perechen()
{

    Part("Открытие документа : " + FileDir);

    WordApplication = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    Part("[Load] Открытие документа : " + FileDir);

    //  WordApplication->setProperty("Visible", 1); //Показать (Открыть) окно MSWord с документом

    WordDocuments = WordApplication->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    Part("[Documents()] Открытие документа : " + FileDir);

    WordDocuments->querySubObject( "Open(%T)",FileDir); //D:\\11111\\One.docx

    Part("[OK] Открытие документа : " + FileDir);

    ActiveDocument = WordApplication->querySubObject("ActiveDocument()"); // Сделать документ активным



    selection2 = WordApplication->querySubObject("Selection()");  // Создать класс Области страницы


    Tables = selection2->querySubObject("Tables(1)"); // Выбираем 1 таблицу в документе


    StartCell  = Tables->querySubObject("Cell(Row, Column)", 6, 2); // (ячейка C1)

    CellRange = StartCell->querySubObject("Range()"); // Область выбранной ячейки

    columns = Tables->querySubObject("Columns")->property("Count").toInt();

    rows = Tables->querySubObject("Rows")->property("Count").toInt();

    qDebug () << "Колонки = " << columns;

    qDebug () <<"Строки = " << rows;


    Part("Открыт : " + FileDir + " Количество Колонок: " + QString::number(columns) + " Строк: " +  QString::number(rows));


    Findelements_Perechen();

}

////////////////////////////////////////////////////////////////

void MYWORD::Findelements_Perechen()
{

    R.clear();      //отчистка резисторы
    RName.clear();  //отчистка имя резисторов


    C_Z.clear();    //отчистка конденсаторы и фильтры
    C_ZName.clear();  //отчистка имя конденсаторов

    XP_XS_XW_X.clear();  //отчистка Вилка
    XP_XS_XW_XName.clear(); //отчистка ИмяВилки

    DA.clear();
    DAName.clear();

    DD.clear();
    DDName.clear();

    BQ_G.clear();
    BQ_GName.clear();

    L.clear();
    LName.clear();

    TV.clear();
    TVName.clear();

    HL_VD.clear();
    HL_VDName.clear();

    VT.clear();
    VTName.clear();


    int count_find = 0;

    QString text;

    emit changeWork(rows);


    Part("Ищем Элементы... Колонок: " + QString::number(columns) + " Строк: " +  QString::number(rows));


    for(int i=1; i <  rows;i++)
    {
        emit changeWork(rows);

        for(int j=1; j < columns; j++)
        {

            StartCell = Tables->querySubObject("Cell(Row, Column)", i, j);
            CellRange = StartCell->querySubObject("Range()");
            text =  CellRange->property("Text").toString();


            int countStart = 0;
            int countCount = 0;

            QString str = "";
            QString countStartString = "";

            for(int k=0; k < text.count();k++)
            {
                if(text[k] == '-')
                {
                    str +=  text[0];

                    if(str == 'R' || str == 'C' || str == 'Z' || (str == 'X' && (text[1] != 'P' && text[1] !='S' && text[1] !='W')))
                    {
                        for(int l=1;l < k;l++)
                        {
                            countStartString += text[l];
                        }

                        countStart = countStartString.toInt();

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+2);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        countCount = text.toInt();

                    }
                    else
                    {
                        str +=  text[1];

                        for(int l=2;l < k;l++)
                        {
                            countStartString += text[l];
                        }

                        countStart = countStartString.toInt();

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+2);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        text.remove(text.count()-2,2);

                        countCount = text.toInt();

                    }

                    break;
                }
            }

            StartCell = Tables->querySubObject("Cell(Row, Column)", i, j);
            CellRange = StartCell->querySubObject("Range()");
            text =  CellRange->property("Text").toString();



            //Ищем R (резисторы)
            if((text[0] == "R") && (j == 2))
            {

                if(str == "R")
                {
                    for(int col = countStart; col <= (countStart+countCount);col++)
                    {
                        count_find++;

                        R.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            R.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);

                                CellRange = StartCell->querySubObject("Range()");

                                text+=CellRange->property("Text").toString();

                                RName.append(text);
                            }
                            else
                            {
                                RName.append(text);
                            }
                        }


                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    count_find++;
                    R.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        R.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            RName.append(text);
                        }
                        else
                        {
                            RName.append(text);
                        }
                    }
                }

                break;

            }

            //Ищем C (конденсаторы)
            if((text[0] == "C") && (j == 2)) //С
            {

                if(str[0] == "C")
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        C_Z.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            C_Z.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);

                                CellRange = StartCell->querySubObject("Range()");

                                text+=CellRange->property("Text").toString();

                                C_ZName.append(text);
                            }
                            else
                            {
                                C_ZName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    C_Z.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        C_Z.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            C_ZName.append(text);
                        }
                        else
                        {
                            C_ZName.append(text);
                        }
                    }
                }


                break;
            }

            //Ищем Z (фильтры)
            if((text[0] == "Z") && (j == 2))
            {

                if(str[0] == "Z")
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        C_Z.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            C_Z.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);

                                CellRange = StartCell->querySubObject("Range()");

                                text+=CellRange->property("Text").toString();

                                C_ZName.append(text);
                            }
                            else
                            {
                                C_ZName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {

                    C_Z.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        C_Z.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            C_ZName.append(text);
                        }
                        else
                        {
                            C_ZName.append(text);
                        }
                    }
                }


                break;
            }


            //Ищем XP (вилка)
            if(((text[0] == "X") && (text[1] == "P")) && (j == 2))
            {

                if(((str[0] == "X") && (str[1] == "P")))
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        XP_XS_XW_X.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            XP_XS_XW_X.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);

                                CellRange = StartCell->querySubObject("Range()");

                                text+=CellRange->property("Text").toString();

                                XP_XS_XW_XName.append(text);
                            }
                            else
                            {
                                XP_XS_XW_XName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {

                    XP_XS_XW_X.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        XP_XS_XW_X.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            XP_XS_XW_XName.append(text);
                        }
                        else
                        {
                            XP_XS_XW_XName.append(text);
                        }
                    }
                }


                break;
            }

            //Ищем XS (Розетка)
            if(((text[0] == "X") && (text[1] == "S")) && (j == 2))
            {

                if(((str[0] == "X") && (str[1] == "S")))
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        XP_XS_XW_X.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            XP_XS_XW_X.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);

                                CellRange = StartCell->querySubObject("Range()");

                                text+=CellRange->property("Text").toString();

                                XP_XS_XW_XName.append(text);
                            }
                            else
                            {
                                XP_XS_XW_XName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    XP_XS_XW_X.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        XP_XS_XW_X.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            XP_XS_XW_XName.append(text);
                        }
                        else
                        {
                            XP_XS_XW_XName.append(text);
                        }
                    }
                }


                break;
            }

            //Ищем XW (Вилка)
            if(((text[0] == "X") && (text[1] == "W")) && (j == 2))
            {

                if(((str[0] == "X") && (str[1] == "W")))
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        XP_XS_XW_X.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            XP_XS_XW_X.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);

                                CellRange = StartCell->querySubObject("Range()");

                                text+=CellRange->property("Text").toString();

                                XP_XS_XW_XName.append(text);
                            }
                            else
                            {
                                XP_XS_XW_XName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {

                    XP_XS_XW_X.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        XP_XS_XW_X.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            XP_XS_XW_XName.append(text);
                        }
                        else
                        {
                            XP_XS_XW_XName.append(text);
                        }
                    }
                }

                break;
            }

            //Ищем X (вилка)
            if((text[0] == "X") && (j == 2))
            {

                if(str[0] == "X")
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        XP_XS_XW_X.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            XP_XS_XW_X.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);

                                CellRange = StartCell->querySubObject("Range()");

                                text+=CellRange->property("Text").toString();

                                XP_XS_XW_XName.append(text);
                            }
                            else
                            {
                                XP_XS_XW_XName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    XP_XS_XW_X.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        XP_XS_XW_X.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            XP_XS_XW_XName.append(text);
                        }
                        else
                        {
                            XP_XS_XW_XName.append(text);

                        }
                    }
                }

                break;
            }


            //Ищем BQ (Резонатор)
            if(((text[0] == "B") && (text[1] == "Q")) && (j == 2))
            {
                if(((str[0] == "B") && (str[1] == "Q")))
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        BQ_G.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            BQ_G.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                                CellRange = StartCell->querySubObject("Range()");
                                text+=CellRange->property("Text").toString();

                                BQ_GName.append(text);
                            }
                            else
                            {
                                BQ_GName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {

                    BQ_G.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    qDebug () << text;

                    if(text == "Отсутствует\r\u0007")
                    {
                        BQ_G.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            BQ_GName.append(text);
                        }
                        else
                        {
                            BQ_GName.append(text);
                        }
                    }
                }

                break;
            }


            //Ищем G
            if((text[0] == "G") && (j == 2))
            {
                if(str[0] == "G")
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        BQ_G.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            BQ_G.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                                CellRange = StartCell->querySubObject("Range()");
                                text+=CellRange->property("Text").toString();

                                BQ_GName.append(text);
                            }
                            else
                            {
                                BQ_GName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    BQ_G.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    qDebug () << text;

                    if(text == "Отсутствует\r\u0007")
                    {
                        BQ_G.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            BQ_GName.append(text);
                        }
                        else
                        {
                            BQ_GName.append(text);
                        }
                    }
                }

                break;
            }

            //Ищем DA (Микросхема)
            if(((text[0] == "D") && (text[1] == "A")) && (j == 2))
            {
                if((str[0] == "D" && str[1] == "A"))
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        DA.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            DA.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                                CellRange = StartCell->querySubObject("Range()");
                                text+=CellRange->property("Text").toString();

                                DAName.append(text);
                            }
                            else
                            {
                                DAName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {

                    DA.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        DA.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            DAName.append(text);
                        }
                        else
                        {
                            DAName.append(text);
                        }
                    }
                }

                break;
            }

            //Ищем DD (Микросхема)
            if(((text[0] == "D") && (text[1] == "D")) && (j == 2))
            {
                if((str[0] == "D" && str[1] == "D"))
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        DD.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            DD.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                                CellRange = StartCell->querySubObject("Range()");
                                text+=CellRange->property("Text").toString();

                                DDName.append(text);
                            }
                            else
                            {
                                DDName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    DD.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        DD.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            DDName.append(text);
                        }
                        else
                        {
                            DDName.append(text);
                        }
                    }
                }

                break;
            }

            //Ищем U (источники питания)
            if((text[0] == "U") && (j == 2))
            {
                if(str[0] == "U")
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        U.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            U.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                                CellRange = StartCell->querySubObject("Range()");
                                text+=CellRange->property("Text").toString();

                                UName.append(text);
                            }
                            else
                            {
                                UName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    U.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();


                    if(text == "Отсутствует\r\u0007")
                    {
                        U.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            UName.append(text);
                        }
                        else
                        {
                            UName.append(text);
                        }
                    }
                }

                break;

            }

            //Ищем L
            if((text[0] == "L") && (j == 2))
            {
                if(str[0] == "L")
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        L.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            L.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                                CellRange = StartCell->querySubObject("Range()");
                                text+=CellRange->property("Text").toString();

                                LName.append(text);
                            }
                            else
                            {
                                LName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    L.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        L.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            LName.append(text);
                        }
                        else
                        {
                            LName.append(text);
                        }
                    }
                }


                break;

            }

            //Ищем TV
            if(((text[0] == "T") && (text[1] == "V"))  && (j == 2))
            {
                if((str[0] == "T") && (str[1] == "V"))
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        TV.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            TV.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                                CellRange = StartCell->querySubObject("Range()");
                                text+=CellRange->property("Text").toString();

                                TVName.append(text);
                            }
                            else
                            {
                                TVName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    TV.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        TV.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            TVName.append(text);
                        }
                        else
                        {
                            TVName.append(text);
                        }
                    }
                }

                break;

            }

            //Ищем VT
            if(((text[0] == "V") && (text[1] == "T"))  && (j == 2))
            {
                if((str[0] == "V") && (str[1] == "T"))
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        VT.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            VT.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                                CellRange = StartCell->querySubObject("Range()");
                                text+=CellRange->property("Text").toString();

                                VTName.append(text);
                            }
                            else
                            {
                                VTName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    VT.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        VT.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            VTName.append(text);
                        }
                        else
                        {
                            VTName.append(text);
                        }
                    }
                }


                break;

            }

            //Ищем HL
            if(((text[0] == "H") && (text[1] == "L"))  && (j == 2))
            {
                if((str[0] == "H") && (str[1] == "L"))
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        HL_VD.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            HL_VD.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                                CellRange = StartCell->querySubObject("Range()");
                                text+=CellRange->property("Text").toString();

                                HL_VDName.append(text);
                            }
                            else
                            {
                                HL_VDName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    HL_VD.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        HL_VD.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                            CellRange = StartCell->querySubObject("Range()");
                            text+=CellRange->property("Text").toString();

                            HL_VDName.append(text);
                        }
                        else
                        {
                            HL_VDName.append(text);
                        }
                    }
                }


                break;

            }

            //Ищем VD
            if(((text[0] == "V") && (text[1] == "D"))  && (j == 2))
            {
                if((str[0] == "V") && (str[1] == "D"))
                {
                    for(int col = countStart; col < (countStart+countCount);col++)
                    {
                        count_find++;

                        HL_VD.append(str+QString::number(col)+"\r\u0007");

                        StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                        CellRange = StartCell->querySubObject("Range()");
                        text =  CellRange->property("Text").toString();

                        if(text == "Отсутствует\r\u0007")
                        {
                            HL_VD.removeLast();
                        }
                        else
                        {
                            if(findRussianLanguage(text))
                            {
                                StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);
                                CellRange = StartCell->querySubObject("Range()");
                                text+=CellRange->property("Text").toString();

                                HL_VDName.append(text);
                            }
                            else
                            {
                                HL_VDName.append(text);
                            }
                        }
                    }

                    emit changeWork(rows);
                    i++;

                    break;
                }
                else
                {
                    HL_VD.append(text);

                    StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);
                    CellRange = StartCell->querySubObject("Range()");
                    text =  CellRange->property("Text").toString();

                    if(text == "Отсутствует\r\u0007")
                    {
                        HL_VD.removeLast();
                    }
                    else
                    {
                        if(findRussianLanguage(text))
                        {
                            StartCell = Tables->querySubObject("Cell(Row, Column)", i+1, j+1);

                            CellRange = StartCell->querySubObject("Range()");

                            text+=CellRange->property("Text").toString();

                            HL_VDName.append(text);
                        }
                        else
                        {
                            HL_VDName.append(text);
                        }
                    }
                }

                break;

            }




        }
    }

    Part("Поиск завершен. Закрытие документа.");

    emit findData(R.count(),C_Z.count(),XP_XS_XW_X.count(),BQ_G.count(),DD.count(),DA.count(),U.count(),L.count(),TV.count(),HL_VD.count(),VT.count());

    //ActiveDocument->dynamicCall("Close (boolean)", false);

    WordApplication->dynamicCall("Quit (void)");


    delete WordApplication;


    // creatShablon();

}

bool MYWORD::findRussianLanguage(QString text)
{
    QString russian = "ЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮЁ";

    if(text[0] == " ")
    {
        text.remove(0,1);
    }



    auto list = text.split(' ');


    for(int i=0; i < text.count()-1;i++)
    {
        if(text[i] == "Т" && text[i+1] =="У")
        {
            return false;
        }

        if(text[i] == "У" && text[i+1] =="Э")
        {
            return false;
        }
    }

    QString str = text.split(' ').last();

    if(str == "ТУ" || str == "УЭ")
    {
        return false;
    }
    else
    {
        str = text.split(' ').value(1).toUpper();



        /*
        if(str == "КОНДЕНСАТОР")
        {
          str = text.split(' ').value(2).toUpper();
        }
        if(str == "МИКРОСХЕМА")
        {
            str = text.split(' ').value(2).toUpper();
        }s

        */
    }

    qDebug () << str;
    qDebug () << text;

    for(int i=0; i < str.count();i++)
    {
        for(int j=0; j < russian.count();j++)
        {
            if(str[i] == russian[j])
            {
                return true;
            }
        }
    }

    return false;
}

void MYWORD::creatShablon()
{
    Part("Создание шаблона с R. ["+saveDir+"//R]");

    if(R.count() > 0)
    {
        CreatShablon_R();
    }

    //this->thread()->msleep(10);

    Part("Создание шаблона с С Z. ["+saveDir+"//СZ]");

    if(C_Z.count() > 0)
    {
        CreatShablon_C_Z();
    }

    //this->thread()->msleep(10);

    Part("Создание шаблона с XP XS XW X. ["+saveDir+"//XPXSXWX]");

    if(XP_XS_XW_X.count() > 0)
    {
        CreatShablon_XP_XS_XW_X();
    }

    //this->thread()->msleep(10);

    Part("Создание шаблона с BQ. ["+saveDir+"//BQ]");

    if(BQ_G.count() > 0)
    {
        CreatShablon_BQ();
    }

    //this->thread()->msleep(10);

    Part("Создание шаблона с DD. ["+saveDir+"//DD]");

    if(DD.count() > 0)
    {
        CreatShablon_DD();
    }

    //this->thread()->msleep(10);

    Part("Создание шаблона с DA. ["+saveDir+"//DA]");

    if(DA.count() > 0)
    {
        CreatShablon_DA();
    }

    //this->thread()->msleep(10);


    Part("Создание шаблона с U. ["+saveDir+"//U]");

    if(U.count() > 0)
    {
        CreatShablon_U();
    }

    //this->thread()->msleep(10);

    Part("Создание шаблона с L. ["+saveDir+"//L]");

    if(L.count() > 0)
    {
        CreatShablon_L();
    }

    //this->thread()->msleep(10);

    Part("Создание шаблона с TV. ["+saveDir+"//TV]");

    if(TV.count() > 0)
    {
        CreatShablon_TV();
    }

    //this->thread()->msleep(10);

    Part("Создание шаблона с HL VD. ["+saveDir+"//HLVD]");

    if(HL_VD.count() > 0)
    {
        CreatShablon_HL_VD();
    }

    //this->thread()->msleep(10);

    Part("Создание шаблона с VT. ["+saveDir+"//VT]");

    if(VT.count() > 0)
    {
        CreatShablon_VT();
    }

    //this->thread()->msleep(10);


    R.clear();      //отчистка резисторы
    RName.clear();  //отчистка имя резисторов


    C_Z.clear();    //отчистка конденсаторы и фильтры
    C_ZName.clear();  //отчистка имя конденсаторов

    XP_XS_XW_X.clear();  //отчистка Вилка
    XP_XS_XW_XName.clear(); //отчистка ИмяВилки

    DA.clear();
    DAName.clear();

    DD.clear();
    DDName.clear();

    BQ_G.clear();
    BQ_GName.clear();

    U.clear();
    UName.clear();

    L.clear();
    LName.clear();

    VT.clear();
    VTName.clear();

    HL_VD.clear();
    HL_VDName.clear();

    Part("Шаблоны созданны.");


}


QString MYWORD::addData_R_Power_NTD(int i)
{
    QString code = "";
    int findIndex = 0;
    QString str = RName[i].split(' ').last();

    if(str[0] == 'C' && str[1] == 'R')
    {
        code += str[2];
        code += str[3];
        code += str[4];
        code += str[5];

        findIndex = r_cr_code.indexOf(code);

        if(findIndex != -1)
        {
            return r_cr_power.value(findIndex);
        }
    }

    return "";

}

QString MYWORD::addData_R_TemperatureRange_NTD(int i)
{
    QString code = "";
    int findIndex = 0;
    QString str = RName[i].split(' ').last();

    if(str[0] == 'C' && str[1] == 'R')
    {
        code += str[2];
        code += str[3];
        code += str[4];
        code += str[5];

        findIndex = r_cr_code.indexOf(code);

        if(findIndex != -1)
        {
            return r_cr_TemperatureRange.value(findIndex);
        }
    }

    return "";

}

QString MYWORD::addData_R_U_NTD(int i,double p)
{
    QString code = "";
    QString str = RName[i].split(' ').last();
    QString strR = str.split('-').last();

    double u = 0;


    if(str[0] == 'C' && str[1] == 'R')
    {
        code += strR[0];

        if(strR[1] == 'R')
        {
            code += ".";
            code += strR[2];
            code += strR[strR.count()-3];

            u = qSqrt(code.toDouble() *p);
        }
        else
        {
            if(strR[2] == 'R')
            {
                code += strR[1];
                code += ".";
                code += strR[strR.count()-3];

                u = qSqrt(code.toDouble() *p);
            }
            else
            {
                code += strR[1];
                code += strR[2];

                u = qSqrt((code.toInt() * qPow(10,QString(strR[strR.count()-3]).toInt()))*p);
            }
        }

        if( u == 0.0 )
        {
            code = "";
            code += str[2];
            code += str[3];
            code += str[4];
            code += str[5];

            int  findIndex = r_cr_code.indexOf(code);

            if(findIndex != -1)
            {
                return r_cr_Void.value(findIndex);
            }
        }
        else
        {
            code = "";
            code += str[2];
            code += str[3];
            code += str[4];
            code += str[5];

            int  findIndex = r_cr_code.indexOf(code);

            if(findIndex != -1)
            {
                if(u > r_cr_Void.value(findIndex).toDouble())
                {
                    return r_cr_Void.value(findIndex);
                }
            }


        }

        return QString::number(u,'f', 2);
    }

    return "";
}


void MYWORD::CreatShablon_R()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",listMYWORD[0]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");

    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");


    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//R");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//R");




    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");




    qDebug () << "Example R: " << (R.count()%3) << " ; " <<  (R.count()/3)+1 << " ; " << (R.count()/3);


    if((R.count()%3) > 0 )
    {
        for(int i=1; i < (R.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateR(QString::number(i));
        }
    }
    else
    {
        for(int i=1; i < (R.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateR(QString::number(i));
        }
    }


    /////////////////////////////////////////////////////


    QAxObject* Tables_2= nullptr,*StartCell_2= nullptr,*CellRange_2= nullptr,*StartCell_2_3= nullptr,*CellRange_2_3= nullptr;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < R.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateR(QString::number(i+1));

        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", R[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", RName[i]);


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 4);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));


            //////////////////////
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 16, 4);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)",addData_R_Power_NTD(i));

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 5);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_R_TemperatureRange_NTD(i));


            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 5);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_R_U_NTD(i,addData_R_Power_NTD(i).toDouble()));


            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 6);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));

            //////////////////////
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 16, 6);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_R_Power_NTD(i));

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 7);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_R_TemperatureRange_NTD(i));

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 7);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_R_U_NTD(i,addData_R_Power_NTD(i).toDouble()));

            break;
        }
        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 8);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));

            //////////////////////
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 16, 8);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_R_Power_NTD(i));

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 9);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_R_TemperatureRange_NTD(i));

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 9);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_R_U_NTD(i,addData_R_Power_NTD(i).toDouble()));

            break;
        }

        }




        if(flag == 3)
        {
            flag =0;

            k++;


            if((R.count()%3) > 0  && k > (R.count()/3)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((R.count()%3) == 0  && k > (R.count()/3))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }

        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//R" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//R");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//R");


    //  ActiveDocument_2->dynamicCall("Close (boolean)", false);
    if(stateViewWord == false)
        WordApplication_2->dynamicCall("Quit (void)");


    delete WordApplication_2;
}

QString MYWORD::addData_C_Power_NTD(int i)
{
    QString codePower = "";
    int findIndex = 0;
    QString str = C_ZName[i].split(' ').last();

    if(str[0] == 'G' && str[1] == 'R' && str[2] == 'M') //GRM155R61A474KE15
    {
        codePower += str[8];
        codePower += str[9];

        findIndex = c_grm_codePower.indexOf(codePower);

        if(findIndex != -1)
        {
            return c_grm_power.value(findIndex);
        }
    }

    if(str[0] == 'N' && str[1] == 'F' && str[2] == 'M')
    {
        codePower += str[11];
        codePower += str[12];

        findIndex = z_nfm_code.indexOf(codePower);

        if(findIndex != -1)
        {
            return z_nfm_power.value(findIndex);
        }
    }

    if(str[0] == 'A' && str[1] == 'V' && str[2] == 'X')
    {
        codePower += str[11];
        codePower += str[12];
        codePower += str[13];

        findIndex = c_avx_codePower.indexOf(codePower);

        if(findIndex != -1)
        {
            return c_avx_power.value(findIndex);
        }
    }

    return "";
}

QString MYWORD::addData_C_TemperatureRange_NTD(int i)
{
    QString codeTemperatureRange = "";
    int findIndex = 0;
    QString str = C_ZName[i].split(' ').last();

    if(str[0] == 'G' && str[1] == 'R' && str[2] == 'M') //GRM155R61A474KE15
    {
        codeTemperatureRange += str[6];
        codeTemperatureRange += str[7];

        findIndex = c_grm_codeTemperatureRange.indexOf(codeTemperatureRange);

        if(findIndex != -1)
        {
            return c_grm_TemperatureRange.value(findIndex);
        }
    }

    if(str[0] == 'A' && str[1] == 'V' && str[2] == 'X')
    {

        return c_avx_TemperatureRange.first();

    }

    return "";
}

////////////////////////////////////////////////////


void MYWORD::CreatShablon_C_Z()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject("Documents()"); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",listMYWORD[2]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



    // ActiveDocument_2->querySubObject("Range()")->dynamicCall("Copy()");


    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");


    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//CZ");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//CZ");





    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");




    qDebug () << "Example C_Z: " << (C_Z.count()%3) << " ; " <<  (C_Z.count()/3)+1 << " ; " << (C_Z.count()/3);

    if((C_Z.count()%3) > 0 )
    {
        for(int i=1; i < (C_Z.count()/3)+1;i++)
        {
            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");


            emit updateC_Z(QString::number(i));
        }
    }
    else
    {
        for(int i=1; i < (C_Z.count()/3);i++)
        {
            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateC_Z(QString::number(i));
        }
    }


    /////////////////////////////////////////////////////


    QAxObject* Tables_2= nullptr,*StartCell_2= nullptr,*CellRange_2= nullptr,*StartCell_2_3= nullptr,*CellRange_2_3= nullptr;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < C_Z.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateC_Z(QString::number(i));

        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", C_Z[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", C_ZName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 4); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 5);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_C_Power_NTD(i));

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 5);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_C_TemperatureRange_NTD(i));

            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 6); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));


            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 7);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_C_Power_NTD(i));

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 7);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_C_TemperatureRange_NTD(i));

            break;
        }
        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 8); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));


            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 9);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_C_Power_NTD(i));

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 9);
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
            CellRange_2_3->dynamicCall("InsertAfter(Text)", addData_C_TemperatureRange_NTD(i));

            break;
        }

        }

        if(flag == 3)
        {
            flag =0;

            k++;

            if((C_Z.count()%3) > 0  && k > (C_Z.count()/3)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((C_Z.count()%3) == 0  && k > (C_Z.count()/3))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }
        }

    }





    //Сохранить pdf
    //  ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//CZ" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//CZ");//"D://11111//1//CZ");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//CZ");//"D://11111//1//CZ");



    //  ActiveDocument_2->dynamicCall("Close (boolean)", false);
    if(stateViewWord == false)
        WordApplication_2->dynamicCall("Quit (void)");

    delete WordApplication_2;
}

void MYWORD::CreatShablon_XP_XS_XW_X()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",listMYWORD[1]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



    // ActiveDocument_2->querySubObject("Range()")->dynamicCall("Copy()");


    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");


    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//XPXSXWX");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//XPXSXWX");




    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");




    qDebug () << "Example XP_XS_XW_X: " << (XP_XS_XW_X.count()%3) << " ; " <<  (XP_XS_XW_X.count()/3)+1 << " ; " << (XP_XS_XW_X.count()/3);


    if((XP_XS_XW_X.count()%3) > 0 )
    {
        for(int i=1; i < (XP_XS_XW_X.count()/3)+1;i++)
        {
            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateXP_XS_XW_X(QString::number(i));
        }
    }
    else
    {
        for(int i=1; i < (XP_XS_XW_X.count()/3);i++)
        {
            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateXP_XS_XW_X(QString::number(i));

        }
    }



    /////////////////////////////////////////////////////


    QAxObject* Tables_2= nullptr,*StartCell_2= nullptr,*CellRange_2= nullptr,*StartCell_2_3= nullptr,*CellRange_2_3= nullptr;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < XP_XS_XW_X.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateXP_XS_XW_X(QString::number(i));
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", XP_XS_XW_X[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", XP_XS_XW_XName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 5); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 7); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }

        }


        if(flag == 3)
        {
            flag =0;

            k++;

            if((XP_XS_XW_X.count()%3) > 0  && k > (XP_XS_XW_X.count()/3)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((XP_XS_XW_X.count()%3) == 0  && k > (XP_XS_XW_X.count()/3))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }
        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//XPXSXW" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//XPXSXWX");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//XPXSXWX");


    //  ActiveDocument_2->dynamicCall("Close (boolean)", false);
    if(stateViewWord == false)
        WordApplication_2->dynamicCall("Quit (void)");

    delete WordApplication_2;
}

void MYWORD::CreatShablon_VT()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",listMYWORD[9]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");

    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//VT");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//VT");


    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");


    qDebug () << "Example VT: " << (VT.count()%3) << " ; " <<  (VT.count()/3)+1 << " ; " << (VT.count()/3);
    if((VT.count()%3) > 0 )
    {
        for(int i=1; i < (VT.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateVT(QString::number(i));
        }
    }
    else
    {
        for(int i=1; i < (VT.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateVT(QString::number(i));

        }
    }



    /////////////////////////////////////////////////////


    QAxObject* Tables_2= nullptr,*StartCell_2= nullptr,*CellRange_2= nullptr,*StartCell_2_3= nullptr,*CellRange_2_3= nullptr;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < VT.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateVT(QString::number(i));
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", VT[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", VTName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 22, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 22, 5); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 22, 7); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }

        }



        if(flag == 3)
        {
            flag =0;

            k++;


            if((VT.count()%3) > 0  && k > (VT.count()/3)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((VT.count()%3) == 0  && k > (VT.count()/3))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }

        }

    }


    //Сохранить pdf

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//VT");//"D://11111//1//BQ");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//VT");//"D://11111//1//BQ");



    //  ActiveDocument_2->dynamicCall("Close (boolean)", false);
    if(stateViewWord == false)
        WordApplication_2->dynamicCall("Quit (void)");

    delete WordApplication_2;
}

void MYWORD::CreatShablon_HL_VD()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",listMYWORD[10]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");

    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//HLVD");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//HLVD");


    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");


    qDebug () << "Example HL_VD: " << (HL_VD.count()%3) << " ; " <<  (HL_VD.count()/3)+1 << " ; " << (HL_VD.count()/3);

    if((HL_VD.count()%3) > 0 )
    {
        for(int i=1; i < (HL_VD.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateHL_VD(QString::number(i));
        }
    }
    else
    {
        for(int i=1; i < (HL_VD.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateHL_VD(QString::number(i));
        }
    }



    /////////////////////////////////////////////////////


    QAxObject* Tables_2= nullptr,*StartCell_2= nullptr,*CellRange_2= nullptr,*StartCell_2_3= nullptr,*CellRange_2_3= nullptr;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < HL_VD.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateHL_VD(QString::number(i));
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", HL_VD[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", HL_VDName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 5); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 7); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }

        }



        if(flag == 3)
        {
            flag =0;

            k++;

            if((HL_VD.count()%3) > 0  && k > (HL_VD.count()/3)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((HL_VD.count()%3) == 0  && k > (HL_VD.count()/3))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }

        }

    }





    //Сохранить pdf

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//HLVD");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//HLVD");



    //  ActiveDocument_2->dynamicCall("Close (boolean)", false);
    if(stateViewWord == false)
        WordApplication_2->dynamicCall("Quit (void)");


    delete WordApplication_2;
}

void MYWORD::CreatShablon_BQ()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",listMYWORD[2]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");

    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//BQG");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//BQG");


    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");


    qDebug () << "Example BQ_G: " << (BQ_G.count()%3) << " ; " <<  (BQ_G.count()/3)+1 << " ; " << (BQ_G.count()/3);

    if((BQ_G.count()%3) > 0 )
    {
        for(int i=1; i < (BQ_G.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateBQ_G(QString::number(i));

        }
    }
    else
    {
        for(int i=1; i < (BQ_G.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateBQ_G(QString::number(i));

        }
    }



    /////////////////////////////////////////////////////


    QAxObject* Tables_2= nullptr,*StartCell_2= nullptr,*CellRange_2= nullptr,*StartCell_2_3= nullptr,*CellRange_2_3= nullptr;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < BQ_G.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateBQ_G(QString::number(i));
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");
        CellRange_2->dynamicCall("InsertAfter(Text)", BQ_G[i]);

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");
        CellRange_2_3->dynamicCall("InsertAfter(Text)", BQ_GName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 5); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 7); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }

        }



        if(flag == 3)
        {
            flag =0;

            k++;

            if((BQ_G.count()%3) > 0  && k > (BQ_G.count()/3)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((BQ_G.count()%3) == 0  && k > (BQ_G.count()/3))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }

        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//BQ" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//BQG");//"D://11111//1//BQ");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//BQG");//"D://11111//1//BQ");



    //  ActiveDocument_2->dynamicCall("Close (boolean)", false);
    if(stateViewWord == false)
        WordApplication_2->dynamicCall("Quit (void)");

    delete WordApplication_2;
}

void MYWORD::CreatShablon_DA()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",listMYWORD[4]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");

    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//DA");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//DA");

    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");


    qDebug () << "Example DA: " << (DA.count()%2) << " ; " <<  (DA.count()/2)+1 << " ; " << (DA.count()/2);

    if((DA.count()%2) > 0 )
    {
        for(int i=1; i < (DA.count()/2)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateDA(QString::number(i));
        }
    }
    else
    {
        for(int i=1; i < (DA.count()/2);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateDA(QString::number(i));

        }
    }






    /////////////////////////////////////////////////////


    QAxObject* Tables_2,*StartCell_2,*CellRange_2,*StartCell_2_3,*CellRange_2_3;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < DA.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateDA(QString::number(i));
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", DA[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", DAName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 4); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 7); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }


        }



        if(flag == 2)
        {
            flag =0;

            k++;


            if((DA.count()%2) > 0  && k > (DA.count()/2)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((DA.count()%2) == 0  && k > (DA.count()/2))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }

        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//DADD" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//DA");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//DA");



    //  ActiveDocument_2->dynamicCall("Close (boolean)", false);
    if(stateViewWord == false)
        WordApplication_2->dynamicCall("Quit (void)");

    delete WordApplication_2;
}

void MYWORD::CreatShablon_DD()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",listMYWORD[5]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");

    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)",saveDir+"//RESULT//DD");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)",saveDir+"//RESULT//DD");

    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");


    qDebug () << "Example DD: " << (DD.count()%2) << " ; " <<  (DD.count()/2)+1 << " ; " << (DD.count()/2);

    if((DD.count()%2) > 0 )
    {
        for(int i=1; i < (DD.count()/2)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateDD(QString::number(i));

        }
    }
    else
    {
        for(int i=1; i < (DD.count()/2);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateDD(QString::number(i));
        }
    }






    /////////////////////////////////////////////////////


    QAxObject* Tables_2= nullptr,*StartCell_2 = nullptr,*CellRange_2= nullptr,*StartCell_2_3= nullptr,*CellRange_2_3= nullptr;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < DD.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateDD(QString::number(i));
        flag++; //чет не так в нижней ячейке

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", DD[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", DDName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 18, 4); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 18, 7); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }


        }



        if(flag == 2)
        {
            flag =0;

            k++;


            if((DD.count()%2) > 0  && k > (DD.count()/2)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((DD.count()%2) == 0  && k > (DD.count()/2))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }

        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//DADD" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)",saveDir+"//RESULT//DD"); //"D://11111//1//DD");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)",saveDir+"//RESULT//DD"); //"D://11111//1//DD");



    //  ActiveDocument_2->dynamicCall("Close (boolean)", false);
    if(stateViewWord == false)
        WordApplication_2->dynamicCall("Quit (void)");

    delete WordApplication_2;
}

void MYWORD::CreatShablon_U()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",listMYWORD[6]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



    // ActiveDocument_2->querySubObject("Range()")->dynamicCall("Copy()");


    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");


    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//U");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//U");




    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");

    qDebug () << "Example U: " << (U.count()%3) << " ; " <<  (U.count()/3)+1 << " ; " << (U.count()/3);

    if((U.count()%3) > 0 )
    {
        for(int i=1; i < (U.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateU(QString::number(i));
        }
    }
    else
    {
        for(int i=1; i < (U.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateU(QString::number(i));

        }
    }

    /////////////////////////////////////////////////////


    QAxObject* Tables_2= nullptr,*StartCell_2= nullptr,*CellRange_2= nullptr,*StartCell_2_3= nullptr,*CellRange_2_3= nullptr;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < U.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateU(QString::number(i));
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", U[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", UName[i]);

        //Темпиратура

        if(flag == 3)
        {
            flag =0;

            k++;


            if((U.count()%3) > 0  && k > (U.count()/3)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((U.count()%3) == 0  && k > (U.count()/3))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }

        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//U" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//U");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//U");



    //  ActiveDocument_2->dynamicCall("Close (boolean)", false);
    if(stateViewWord == false)
        WordApplication_2->dynamicCall("Quit (void)");

    delete WordApplication_2;
}

void MYWORD::CreatShablon_L()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",listMYWORD[7]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



    // ActiveDocument_2->querySubObject("Range()")->dynamicCall("Copy()");


    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//L");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//L");

    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");




    qDebug () << "Example L: " << (L.count()%3) << " ; " <<  (L.count()/3)+1 << " ; " << (L.count()/3);


    if((L.count()%3) > 0 )
    {
        for(int i=1; i < (L.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateL(QString::number(i));

        }
    }
    else
    {
        for(int i=1; i < (L.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateL(QString::number(i));
        }
    }


    /////////////////////////////////////////////////////


    QAxObject* Tables_2= nullptr,*StartCell_2= nullptr,*CellRange_2= nullptr,*StartCell_2_3= nullptr,*CellRange_2_3= nullptr;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < L.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateL(QString::number(i));
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", L[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", LName[i]);

        //Темпиратура

        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 5); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }

        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 7); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }

        }



        if(flag == 3)
        {
            flag =0;

            k++;


            if((L.count()%3) > 0  && k > (L.count()/3)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((L.count()%3) == 0  && k > (L.count()/3))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }

        }

    }







    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//L" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//L");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//L");




    //  ActiveDocument_2->dynamicCall("Close (boolean)", false);
    if(stateViewWord == false)
        WordApplication_2->dynamicCall("Quit (void)");

    delete WordApplication_2;
}

void MYWORD::CreatShablon_TV()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", stateViewWord);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject("Open(%T)",listMYWORD[8]); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");

    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject("Range()")->dynamicCall("Copy()");

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//TV");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//TV");

    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");



    qDebug () << "Example TV: " << (TV.count()%3) << " ; " <<  (TV.count()/3)+1 << " ; " << (TV.count()/3);


    if((TV.count()%3) > 0 )
    {
        for(int i=1; i < (TV.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateTV(QString::number(i));
        }
    }
    else
    {
        for(int i=1; i < (TV.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");
            selection_2->dynamicCall("InsertBreak()");
            //this->thread()->msleep(10);
            selection_2->dynamicCall("Paste()");

            emit updateTV(QString::number(i));

        }
    }


    /////////////////////////////////////////////////////


    QAxObject* Tables_2= nullptr,*StartCell_2= nullptr,*CellRange_2= nullptr,*StartCell_2_3= nullptr,*CellRange_2_3= nullptr;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < TV.count();i++)
    {
        //this->thread()->msleep(10);
        emit updateTV(QString::number(i));
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", TV[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", TVName[i]);

        //Темпиратура

        if(flag == 3)
        {
            flag =0;

            k++;


            if((TV.count()%3) > 0  && k > (TV.count()/3)+1)
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                if((TV.count()%3) == 0  && k > (TV.count()/3))
                {
                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    qDebug () << "K = " << k;
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));
                }
            }

        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//L" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs (const QString&)", saveDir+"//RESULT//TV");
    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", saveDir+"//RESULT//TV");



    //ActiveDocument_2->dynamicCall("Close (boolean)", false);


    if(stateViewWord == false)
    {
        WordApplication_2->dynamicCall("Quit (void)");
    }

    delete WordApplication_2;

}


