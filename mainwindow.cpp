#include "mainwindow.h"
#include "ui_mainwindow.h"



MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);


    word = new MYWORD();

    connect(this,&MainWindow::scanDir,word,&MYWORD::scanDirWork);
    connect(word,&MYWORD::scanning,this,&MainWindow::on_scanningList);

    qRegisterMetaType<QDir>("QDir");

//    QAxWidget *pdf = new QAxWidget();
//    pdf->setControl("Adobe PDF Reader");
//    pdf->dynamicCall("LoadFile(const QString&)", "D:/1/1.pdf");


    QAxWidget* Widget = new QAxWidget();

    Widget->resize(500,500);


    Widget->setWindowTitle("D:/1/1.pdf");


    Widget->setControl(QString::fromUtf8("{8856F961-340A-11D0-A96B-00C04FD705A2}"));
    Widget->dynamicCall("Navigate(const QString&)", QString("D:/1/1.pdf"));

   // Widget->dynamicCall("Find(const QString&)","С");

    Widget->show();
}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_pushButton_clicked()
{

//    QFileDialog dialog(this);
//    dialog.setFileMode(QFileDialog::AnyFile);


    QDir dir;
    dir.setPath(ui->lineEdit->text());

    scanDir(ui->textEdit->document()->toPlainText(), dir);

}

void MainWindow::on_scanningList(QString data, int i, int N)
{
    if(ui->progressBar->maximum() != N)
        ui->progressBar->setMaximum(N);

    ui->progressBar->setValue(i);
    ui->statusBar->showMessage(data);
}


void MainWindow::on_pushButton_2_clicked()
{
    qDebug() << "==============================";
    qDebug() << word->listFiles;

}
