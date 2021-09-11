package com.company;

import java.io.*;
import java.util.Iterator;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xlsxToCSV {

   static void readXlsx(File inputFile, File outputFile) {
       StringBuffer data = new StringBuffer();

       try{
           FileInputStream fis = new FileInputStream(inputFile);
           FileOutputStream fos = new FileOutputStream(outputFile);
//           OutputStreamWriter osw = new OutputStreamWriter(outputFile);

           //read file
           Workbook workbook = null;
           //write file
           Workbook workbook1 = null;

           String ext  = FilenameUtils.getExtension(inputFile.toString());

           if(ext.equalsIgnoreCase("xlsx")) {
               workbook = new XSSFWorkbook(fis);
           }else if(ext.equalsIgnoreCase("xls")){
               workbook = new HSSFWorkbook(fis);
           }
           int numberOfSheets = workbook.getNumberOfSheets();
            Row row;
            Cell cell1, cell2;
            Sheet sheet = workbook.getSheetAt(0);
            String[] titles = new String[] {"login_name","userpassword","name","depid","email"};

           for(String title : titles){
               fos.write(title.getBytes());
               fos.write(",".getBytes());
           }
           fos.write("\r\n".getBytes());

           for(int i = 1; i <= sheet.getLastRowNum(); i++) {
               row = sheet.getRow(i);
//               System.out.println(row.toString());
               cell1 = row.getCell(1);                //teacher's name
               cell2 = row.getCell(6);                //e-mail


               //trim()

               //login_name
               fos.write(cell2.toString().trim().getBytes());
               fos.write("@mail.klcg.gov.tw".getBytes());
               fos.write(",".getBytes());
               //userpassword
               fos.write("password".getBytes());
               fos.write(",".getBytes());
               //name
               fos.write(cell1.toString().getBytes());
               fos.write(",".getBytes());
               //depid
               fos.write("100".getBytes());
               fos.write(",".getBytes());
               //email
               fos.write(cell2.toString().trim().getBytes());
               fos.write("@mail.klcg.gov.tw".getBytes());
               fos.write(",".getBytes());
               fos.write("\r\n".getBytes());
           }

       } catch(Exception e) {
           System.out.println("*****read excel*****");
           e.printStackTrace();
       }
   }

//************************************************************************


    public static void main(String[] args) {
        // reading file from desktop
        File inputFile = new File("/home/user/Downloads/govMail/七堵國小.xlsx"); //provide your path
        // writing excel data to csv
        File outputFile = new File("/home/user/Downloads/govMail/importCSV/七堵國小.csv");  //provide your path
        //work space
        //xlsx(inputFile, outputFile);
        readXlsx(inputFile, outputFile);

        System.out.println("Conversion of " + inputFile + " to flat file: "
                + outputFile + " is completed");
    }

}
