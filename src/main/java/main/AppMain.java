package main;

import form.MainWin;

import javax.swing.*;

//compile cmd:
//java -jar TimeSheetFormat.jar
public class AppMain {


    public static void main(String[] args) {
        MainWin mainWin = new MainWin();
        JFrame jFrame = new JFrame("Timesheet Creating");
        mainWin.setJf(jFrame);
        jFrame.setContentPane(mainWin.getPanel1());
        jFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        jFrame.setSize (1200, 350);
        jFrame.pack();
        jFrame.setResizable(true);
        jFrame.setVisible(true);
    }
}
