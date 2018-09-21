#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QTextBrowser>

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

    qRegisterMetaType<QDir>("QDir");

    //    QAxWidget *pdf = new QAxWidget();
    //    pdf->setControl("Adobe PDF Reader");
    //    pdf->dynamicCall("LoadFile(const QString&)", "D:/1/1.pdf");


//    Widget = new QAxWidget();

//    Widget->resize(500,500);


//    Widget->setWindowTitle("D:/1/1.pdf");



//    QString str = "\"При проведении аудита продажи готовой продукции используются процедуры\"";


//    Widget->setControl(QString::fromUtf8("{8856F961-340A-11D0-A96B-00C04FD705A2}"));
//    Widget->dynamicCall("Navigate(const QString&)", QString("D:\\Work\\ПРОТОКОЛЫ\\Аудит\\Protokol_po_tiestu_[Itoghovyi_tiest._Audit_i_kontrol'._2017_Dvorietskaia_]_na_10_12_2017(1).pdf#navpanes=1=OpenActions&search="+str+"&toolbar=0"));


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
