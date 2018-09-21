#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>


#include <QDir>
#include <QDebug>

#include <word.h>

#include <QFileDialog>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    QFileDialog* dialog;

    MYWORD* word;
    QAxWidget* Widget;

signals:
    void scanDir(QString str,QDir dir);

private slots:
    void on_pushButton_clicked();
    void on_scanningList(QString data,int i,int N);

    void on_pushButton_2_clicked();

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
