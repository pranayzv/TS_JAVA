package main;

import model.EmpInfo;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Locale;

public class TS {
    public  XSSFFont headingFont, mainBoldFont;
    public  XSSFCellStyle headStyle, mainStyle,contentStyle,leftAlignBoldStyle,rightContentStyle,bottomContentStyle,rightBold,ulDate;

    public  ArrayList<Integer> weekoffs = new ArrayList<>();
    public  ArrayList<String> workDesc = new ArrayList<>();

     int dateStartRowCn = 5;
    private  XSSFCellStyle rightBottomStyle,endTop,endMid,endBottom;

    private EmpInfo empInfo;
    private JFrame jFrame;

    public TS(){}

   public TS(EmpInfo empInfo){
        this.empInfo = empInfo;
    }

    public void init() throws Exception{



        XSSFWorkbook wb  = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("Doyen-Timesheet");

        Calendar calendar = Calendar.getInstance();

        calendar.set(Calendar.MONTH, empInfo.getMonth()-1);
        calendar.set(Calendar.YEAR, empInfo.getYear());
        calendar.set(Calendar.DAY_OF_MONTH,1);

        workDesc = getTable();
        workDesc.remove(0);

        intiStyle(wb);
        //intiSpaces(sheet);


        //merge for heading.
        sheet.addMergedRegion(new CellRangeAddress(1,1,0,4));
        Row row = sheet.createRow(1);

        Cell cell = row.createCell(0);
        cell.setCellValue("TIME SHEET");
        headStyle.setFont(headingFont);
        cell.setCellStyle(headStyle);
        cell = row.createCell(1);
        cell.setCellStyle(headStyle);
        cell = row.createCell(2);
        cell.setCellStyle(headStyle);
        cell = row.createCell(3);
        cell.setCellStyle(headStyle);
        cell = row.createCell(4);
        cell.setCellStyle(headStyle);


        //merge for 3 cells.
        sheet.addMergedRegion(new CellRangeAddress(2,2,0,3));

        row = sheet.createRow(2);
        cell = row.createCell(0);
        cell.setCellValue("Name: "+empInfo.getEmpName());
        mainStyle.setFont(mainBoldFont);
        mainStyle.setAlignment(CellStyle.ALIGN_RIGHT);
        cell.setCellStyle(leftAlignBoldStyle);

        createCell(row, 1,"");
        createCell(row, 2,"");
        createCell(row, 3,"");

        cell = row.createCell(4);
        cell.setCellValue("Reporting Period");
        cell.setCellStyle(rightBold);
        mainStyle.setAlignment(CellStyle.ALIGN_CENTER);

        sheet.addMergedRegion(new CellRangeAddress(3,3,0,3));
        row = sheet.createRow(3);
        cell = row.createCell(0);
        cell.setCellValue("Location: "+empInfo.getLoc());
        mainStyle.setFont(mainBoldFont);
        cell.setCellStyle(leftAlignBoldStyle);

        cell = row.createCell(4);
        String mdate = new SimpleDateFormat("EE-MMMM-yyyy", Locale.ENGLISH).format(new SimpleDateFormat("yyyy-MM-dd").parse(String.format("%d-%d-%d", empInfo.getYear(), empInfo.getMonth(), 1)));

        cell.setCellValue("01st "+mdate.split("-")[1]+" "+empInfo.getYear()+" to "+((calendar.getActualMaximum(Calendar.DAY_OF_MONTH)))+"th "+mdate.split("-")[1]+" "+ empInfo.getYear());
        cell.setCellStyle(ulDate);


        //setting columns:

        row = sheet.createRow(4);
        createBoldCell(row,0,"Date");
        createBoldCell(row,1,"Client");
        createBoldCell(row,2,"Time In");
        createBoldCell(row,3,"Time Out");
        cell = row.createCell(4);
        cell.setCellValue("Activity");
        cell.setCellStyle(rightBold);






        for(int i = 1; i<=calendar.getActualMaximum(Calendar.DAY_OF_MONTH);i++) {


             String date = new SimpleDateFormat("EE-MMM-yyyy", Locale.ENGLISH).format(new SimpleDateFormat("yyyy-MM-dd").parse(String.format("%d-%d-%d", empInfo.getYear(), empInfo.getMonth(), i)));
           
            String curDay = date.split("-")[0];

            if (curDay.equals("Sun") || curDay.equals("Sat")) {

                System.out.println("********Day: " + curDay);
                weekoffs.add(i);
                System.out.println(i);
            }
        }

        int wdi =0;
        for(int i = 1; i<=calendar.getActualMaximum(Calendar.DAY_OF_MONTH);i++){

            String date = new SimpleDateFormat("EE-MMM-yyyy", Locale.ENGLISH).format(new SimpleDateFormat("yyyy-MM-dd").parse(String.format("%d-%d-%d", empInfo.getYear(), empInfo.getMonth(), i)));
            row = sheet.createRow(dateStartRowCn);
            String curDay = date.split("-")[0] ;

            createCell(row,0,i+"-"+date.split("-")[1]+"-"+date.split("-")[2]);

            dateStartRowCn++;
        }

        int wo = 0;
        dateStartRowCn = 5;
        int monthDays = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
        for(int i = 1; i<=calendar.getActualMaximum(Calendar.DAY_OF_MONTH);i++) {


             String date = new SimpleDateFormat("EE-MMM-yyyy", Locale.ENGLISH).format(new SimpleDateFormat("yyyy-MM-dd").parse(String.format("%d-%d-%d", empInfo.getYear(), empInfo.getMonth(), i)));
            row = sheet.getRow(dateStartRowCn);
            String curDay = date.split("-")[0];


            if(isWeekOff(i)){
                if(curDay.equals("Sat") && i<monthDays){
                    sheet.addMergedRegion(new CellRangeAddress(dateStartRowCn,dateStartRowCn+1,1,4));
                    createCell(row,2,"");
                    createCell(row,3,"");
                    createCell(row,4,"");
                    createCell(row,1,"Week Off");
                    i++;
                    dateStartRowCn++;

                    row = sheet.getRow(dateStartRowCn);
                    createCell(row,2,"");
                    createCell(row,3,"");
                    createCell(row,4,"");
                    createCell(row,1,"");
                }else{
                    sheet.addMergedRegion(new CellRangeAddress(dateStartRowCn,dateStartRowCn,1,4));
                    createCell(row,2,"");
                    createCell(row,3,"");
                    createCell(row,4,"");
                    createCell(row,1,"Week Off");
                }

            }
            else {

                //do normal entry
                createCell(row,1,"External");
                createCell(row,2,"09:30:00 AM");
                createCell(row,3,"06:30:00 PM");
                if(workDesc.size()>wdi) {
                    createCell(row, 4, workDesc.get(wdi));
                }else {createCell(row, 4, "Activity");}
                wdi++;
            }

            dateStartRowCn++;
        }

        setBorders(sheet,calendar.getActualMaximum(Calendar.DAY_OF_MONTH));
        createEnding(sheet,calendar.getActualMaximum(Calendar.DAY_OF_MONTH));

        sheet.autoSizeColumn(0);
        sheet.setColumnWidth(1, 15 * 256);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.setColumnWidth(4, 60 * 256);
        //outputting the file :

        JFileChooser f = new JFileChooser();
        f.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        f.showOpenDialog(jFrame);

        try (FileOutputStream outputStream = new FileOutputStream(f.getSelectedFile().getPath()+"\\"+empInfo.getEmpName()+"_Monthly_Time_Sheet-"+mdate.split("-")[1]+".xlsx")) {
            wb.write(outputStream);
            outputStream.close();
            msgDiag("File has been successfully saved on disk.");

        }








    }
    public void msgDiag(String msg){
        JOptionPane.showMessageDialog(null,msg);
    }
    public  void setBorders(XSSFSheet sheet,int maxRows){
        maxRows+=4;
        System.out.println("max row:"+maxRows);
        int startRow= 5;
        Row row;
        for (int i = startRow; i <=maxRows ; i++) {
            row = sheet.getRow(i);
            Cell cell = row.getCell(4);
            cell.setCellStyle(rightContentStyle);
        }

        row = sheet.getRow(maxRows);
        Cell cell = row.getCell(0);
        cell.setCellStyle(bottomContentStyle);



        cell = row.getCell(1);
        cell.setCellStyle(bottomContentStyle);



        cell = row.getCell(2);
        cell.setCellStyle(bottomContentStyle);



        cell = row.getCell(3);
        cell.setCellStyle(bottomContentStyle);




        cell = row.getCell(4);
        cell.setCellStyle(rightBottomStyle);
    }


    public  void createEnding(XSSFSheet sheet,int maxrow){

        Cell cell;Row row;
        maxrow+=5;
        maxrow+=1;
        row = sheet.createRow(maxrow);

        cell = row.createCell(0);cell.setCellValue("Prepared By:");cell.setCellStyle(endTop);
        cell = row.createCell(4);cell.setCellValue("Approved By:");cell.setCellStyle(endTop);

        maxrow+=1;
        row = sheet.createRow(maxrow);

        cell = row.createCell(0);cell.setCellValue(empInfo.getEmpName());cell.setCellStyle(endMid);
        cell = row.createCell(4);cell.setCellValue(empInfo.getApproverName());cell.setCellStyle(endMid);

        maxrow+=1;
        row = sheet.createRow(maxrow);

        cell = row.createCell(0);cell.setCellValue("(Employee Name)");cell.setCellStyle(endBottom);
        cell = row.createCell(4);cell.setCellValue("(Reporting Manager Name)");cell.setCellStyle(endBottom);



    }

    public  void intiStyle(XSSFWorkbook wb){

        headingFont = wb.createFont();
        headingFont.setBold(true);
        headingFont.setFontName("Calibri");
        headStyle = wb.createCellStyle();


        rightBold =wb.createCellStyle();
        rightBold.setFont(headingFont);
        rightBold.setAlignment(CellStyle.ALIGN_CENTER);
        rightBold.setBorderTop(CellStyle.BORDER_THIN);
        rightBold.setBorderBottom(CellStyle.BORDER_THIN);
        rightBold.setBorderLeft(CellStyle.BORDER_THIN);
        rightBold.setBorderRight(CellStyle.BORDER_MEDIUM);


        Font ulFont = wb.createFont();
        ulFont.setUnderline(HSSFFont.U_SINGLE);

        ulDate = wb.createCellStyle();
        ulDate.setAlignment(CellStyle.ALIGN_CENTER);
        ulDate.setFont(ulFont);
        ulDate.setBorderTop(CellStyle.BORDER_THIN);
        ulDate.setBorderBottom(CellStyle.BORDER_THIN);
        ulDate.setBorderLeft(CellStyle.BORDER_THIN);
        ulDate.setBorderRight(CellStyle.BORDER_MEDIUM);

        headStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        headStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
        headStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        headStyle.setAlignment(CellStyle.ALIGN_CENTER);
        headStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        headStyle.setFillPattern((short)1);


        endTop = wb.createCellStyle();
        endTop.setAlignment(CellStyle.ALIGN_CENTER);
        endTop.setFont(headingFont);
        endTop.setBorderTop(CellStyle.BORDER_MEDIUM);
        endTop.setBorderBottom(CellStyle.BORDER_THIN);
        endTop.setBorderLeft(CellStyle.BORDER_MEDIUM);
        endTop.setBorderRight(CellStyle.BORDER_MEDIUM);

        endMid = wb.createCellStyle();
        endMid.setAlignment(CellStyle.ALIGN_CENTER);
        endMid.setBorderTop(CellStyle.BORDER_THIN);
        endMid.setBorderBottom(CellStyle.BORDER_THIN);
        endMid.setBorderLeft(CellStyle.BORDER_MEDIUM);
        endMid.setBorderRight(CellStyle.BORDER_MEDIUM);

        endBottom = wb.createCellStyle();
        endBottom.setAlignment(CellStyle.ALIGN_CENTER);
        endBottom.setFont(headingFont);
        endBottom.setBorderTop(CellStyle.BORDER_THIN);
        endBottom.setBorderBottom(CellStyle.BORDER_MEDIUM);
        endBottom.setBorderLeft(CellStyle.BORDER_MEDIUM);
        endBottom.setBorderRight(CellStyle.BORDER_MEDIUM);

        mainBoldFont = wb.createFont();
        mainBoldFont.setBold(true);
        mainBoldFont.setFontName("Calibri");

        contentStyle = wb.createCellStyle();
        contentStyle.setAlignment(CellStyle.ALIGN_CENTER);
        contentStyle.setBorderTop(CellStyle.BORDER_THIN);
        contentStyle.setBorderBottom(CellStyle.BORDER_THIN);
        contentStyle.setBorderLeft(CellStyle.BORDER_THIN);
        contentStyle.setBorderRight(CellStyle.BORDER_THIN);
        //  contentStyle.setWrapText(false);

        rightContentStyle = wb.createCellStyle();
        rightContentStyle.setAlignment(CellStyle.ALIGN_CENTER);
        rightContentStyle.setBorderTop(CellStyle.BORDER_THIN);
        rightContentStyle.setBorderBottom(CellStyle.BORDER_THIN);
        rightContentStyle.setBorderLeft(CellStyle.BORDER_THIN);
        rightContentStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        rightContentStyle.setWrapText(true);

        rightBottomStyle =wb.createCellStyle();
        rightBottomStyle.setAlignment(CellStyle.ALIGN_CENTER);
        rightBottomStyle.setBorderTop(CellStyle.BORDER_THIN);
        rightBottomStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        rightBottomStyle.setBorderLeft(CellStyle.BORDER_THIN);
        rightBottomStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        // rightBottomStyle.setWrapText(true);

        bottomContentStyle = wb.createCellStyle();
        bottomContentStyle.setAlignment(CellStyle.ALIGN_CENTER);
        bottomContentStyle.setBorderTop(CellStyle.BORDER_THIN);
        bottomContentStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        bottomContentStyle.setBorderLeft(CellStyle.BORDER_THIN);
        bottomContentStyle.setBorderRight(CellStyle.BORDER_THIN);
        //bottomContentStyle.setWrapText(true);

        leftAlignBoldStyle = wb.createCellStyle();
        leftAlignBoldStyle.setFont(headingFont);
        leftAlignBoldStyle.setAlignment(CellStyle.ALIGN_LEFT);
        leftAlignBoldStyle.setBorderTop(CellStyle.BORDER_THIN);
        leftAlignBoldStyle.setBorderBottom(CellStyle.BORDER_THIN);
        leftAlignBoldStyle.setBorderLeft(CellStyle.BORDER_THIN);
        leftAlignBoldStyle.setBorderRight(CellStyle.BORDER_THIN);

        mainStyle  = wb.createCellStyle();
        mainStyle.setAlignment(CellStyle.ALIGN_CENTER);
        mainStyle.setBorderTop(CellStyle.BORDER_THIN);
        mainStyle.setBorderBottom(CellStyle.BORDER_THIN);
        mainStyle.setBorderLeft(CellStyle.BORDER_THIN);
        mainStyle.setBorderRight(CellStyle.BORDER_THIN);



    }


    public  boolean isWeekOff(int i){
        for(int d  :weekoffs){
            if(i == d ){
                System.out.println("WEEK OFF DAY: "+i);
                return true;
            }
        }

        return false;

    }


    public  void createBoldCell(Row row, int index, String value){
        Cell cell = row.createCell(index);
        cell.setCellValue(value);
        mainStyle.setFont(mainBoldFont);
        cell.setCellStyle(mainStyle);
    }


    public  void createCell(Row row, int index, String value){
        Cell cell = row.createCell(index);
        cell.setCellValue(value);
        cell.setCellStyle(contentStyle);
    }


    public  ArrayList<String> getTable() {
        ArrayList<String> data = new ArrayList<>();

        try (
                Reader reader = Files.newBufferedReader(Paths.get("C:\\Users\\pranayr\\Desktop\\User_ Pranay Rane_2022-07-01_2022-07-31.csv"));
                CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT);
        ){

            for (CSVRecord record : csvParser) {

                String val = record.get(22);
                System.out.println(val);
                data.add(val);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return data;
    }
    public void setjFrame(JFrame jFrame) {
        this.jFrame = jFrame;
    }
}
