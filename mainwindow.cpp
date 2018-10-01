#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QTextBrowser>
#include <QMessageBox>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);


    dialog = new QFileDialog(this);
    dialog->setFileMode(QFileDialog::Directory);

    word = new MYWORD();

    connect(this,&MainWindow::scanDir,word,&MYWORD::scanDirWork,Qt::QueuedConnection);
    connect(word,&MYWORD::scanning,this,&MainWindow::on_scanningList);
    connect(word,&MYWORD::findWordFinish,this,&MainWindow::findWordFinish);
    connect(this,&MainWindow::closeAllWord,word,&MYWORD::closeAllWord,Qt::QueuedConnection);

    qRegisterMetaType<QDir>("QDir");


    //def = ui->textEdit->document()->toPlainText();

    //    QAxWidget *pdf = new QAxWidget();
    //    pdf->setControl("Adobe PDF Reader");
    //    pdf->dynamicCall("LoadFile(const QString&)", "D:/1/1.pdf");


    //    Widget = new QAxWidget();

    //    Widget->resize(500,500);


    //    Widget->setWindowTitle("D:/1/1.pdf");



    //   QString str = "\"контроль, проводимый по решению государственных органов\"";


    //    Widget->setControl(QString::fromUtf8("{8856F961-340A-11D0-A96B-00C04FD705A2}"));
    //    Widget->dynamicCall("Navigate(const QString&)", QString("D:\\Work\\ПРОТОКОЛЫ\\Аудит\\Protokol_po_tiestu_[Itoghovyi_tiest._Audit_i_kontrol'._2017_Dvorietskaia_]_na_10_12_2017(1).pdf#navpanes=1=OpenActions&toolbar=0&search="+str+""));//


    //    Widget->show();


}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_pushButton_clicked()
{

    QDir dir;
    dir.setPath(ui->lineEdit->text());

    emit scanDir(ui->textEdit->document()->toPlainText(), dir);

    ui->statusBar->showMessage("Старт");
    ui->pushButton->setEnabled(false);
    ui->pushButton_2->setEnabled(false);
    ui->pushButton_3->setEnabled(false);
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
    //qDebug() << word->listFiles;


    if (dialog->exec())
    {

        ui->lineEdit->setText(dialog->directory().path());
        ui->pushButton->setEnabled(true);
    }

}

void MainWindow::findWordFinish()
{

    ui->pushButton->setEnabled(true);
    ui->pushButton_2->setEnabled(true);
    ui->pushButton_3->setEnabled(true);
    ui->statusBar->showMessage("Финиш");

    ui->textEdit->setText("");

}

void MainWindow::on_pushButton_3_clicked()
{
    emit closeAllWord();
    ui->pushButton_3->setEnabled(false);
}
