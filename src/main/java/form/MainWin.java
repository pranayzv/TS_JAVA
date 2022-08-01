package form;

import main.TS;
import model.EmpInfo;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.FocusEvent;
import java.awt.event.FocusListener;
import java.util.Objects;


public class MainWin  {
    private JPanel panel1;
    private JButton openFile1Button;
    private JButton openFile2Button;
    private JButton mergeButton;
    private JTextField textField1;
    private JTextField textField2;
    private JTextField textField4;
    private JComboBox comboBox1;
    private JTextField textField3;
    private JTextField mahapeTextField;
    private JFrame jf;
    private String file1Name="",file2Name="", id1="",id2="";
    String t1msg="",t2msg="";
    String yearmsg = "2022";

    EmpInfo empInfo;

    public MainWin() {
        //setTestData();
        textFieldDesign();
        setListeners();
    }

    public void setListeners(){
        openFile1Button.addActionListener(e -> {
            JFileChooser jFileChooser  = new JFileChooser();
            jFileChooser.showOpenDialog(jf);
            FileNameExtensionFilter filter =  new FileNameExtensionFilter("Excel files", "csv");
            jFileChooser.addChoosableFileFilter(filter);
            jFileChooser.setFileFilter(filter);
            if(jFileChooser.getSelectedFile()==null){msgDiag("Please select a .csv file");return;}
            file1Name = jFileChooser.getSelectedFile().getPath();
            openFile1Button.setText(jFileChooser.getSelectedFile().getName());
        });

        mergeButton.addActionListener(e -> {
            //if(file1Name.isEmpty()){msgDiag("Please select a file. ");return;}
            if(textField4.getText().isEmpty() ){msgDiag("Please enter year. ");return;}
            if(textField2.getText().isEmpty() ){msgDiag("Please enter employee name. ");return;}
            if(textField3.getText().isEmpty() ){msgDiag("Please enter approver's name. ");return;}
            if(mahapeTextField.getText().isEmpty() ){msgDiag("Please enter Location. ");return;}
            if(textField1.getText().isEmpty() ){msgDiag("Please enter client. ");return;}
            
            try {
                int year = Integer.parseInt(textField4.getText());
            }catch (NumberFormatException numberFormatException){
                msgDiag("Please enter year in numbers like : '2022'. ");return;
            }
            if(textField4.getText().length()!=4){
                msgDiag("year length should be 4 like : '2022'. ");return;
            }
            empInfo = new EmpInfo(
                    file1Name,
                    Integer.parseInt(Objects.requireNonNull(comboBox1.getSelectedItem()).toString().split("-")[1]),
                    Integer.parseInt(textField4.getText()),
                    textField2.getText(),
                    textField3.getText(),
                    mahapeTextField.getText(),
                    textField1.getText()
            );
            TS ts = new TS(empInfo);
            ts.setjFrame(jf);

            try{ts.init();}catch (Exception ignored){}

            System.out.println(empInfo);
            
        });
    }

    public void setTestData(){
        file1Name="D:\\Database\\AllExport\\AllXlsx\\d1.xls";
        file2Name="D:\\Database\\AllExport\\AllXlsx\\d2.xls";
        id1="ITM_ID";
        id2="ITM_ID";
    }
public void textFieldDesign(){
    textField1.setForeground(Color.GRAY);
    textField1.addFocusListener(new FocusListener() {
        @Override
        public void focusGained(FocusEvent e) {
            if (textField1.getText().equals(t1msg)) {
                textField1.setText("");
                textField1.setForeground(Color.BLACK);
            }
        }
        @Override
        public void focusLost(FocusEvent e) {
            if (textField1.getText().isEmpty()) {
                textField1.setForeground(Color.GRAY);
                textField1.setText(t2msg);
            }
        }
    });

    textField2.setForeground(Color.GRAY);
    textField2.addFocusListener(new FocusListener() {
        @Override
        public void focusGained(FocusEvent e) {
            if (textField2.getText().equals("")) {
                textField2.setText("");
                textField2.setForeground(Color.BLACK);
            }
        }
        @Override
        public void focusLost(FocusEvent e) {
            if (textField2.getText().isEmpty()) {
                textField2.setForeground(Color.GRAY);
                textField2.setText("");
            }
        }
    });
    if (textField2.getText().isEmpty()) {
        textField2.setForeground(Color.GRAY);
        textField2.setText("");
    }
    if (textField1.getText().isEmpty()) {
        textField1.setForeground(Color.GRAY);
        textField1.setText("");
    }
}
    public JPanel getPanel1() {
        return panel1;
    }

    public void setJf(JFrame jf) {
        this.jf = jf;
    }

    public void msgDiag(String msg){
        JOptionPane.showMessageDialog(null,msg);
    }
}
