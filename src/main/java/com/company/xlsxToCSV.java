package com.company;

import java.io.*;
import java.util.Iterator;
import java.util.Locale;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xlsxToCSV {

   static void readXlsx(File inputFile, File outputFile) {
       StringBuffer data = new StringBuffer();

       try{
           FileInputStream fis = new FileInputStream(inputFile);
           //FileOutputStream fos = new FileOutputStream(outputFile);
           OutputStreamWriter osw = new OutputStreamWriter(new FileOutputStream(outputFile),"UTF-8");
           BufferedWriter bw = new BufferedWriter(osw);
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

           bw.write('\ufeff');
           for(String title : titles){
               bw.write(title);
               bw.write(",");
           }
           bw.write("\r\n");

           for(int i = 1; i <= sheet.getLastRowNum(); i++) {
               row = sheet.getRow(i);
//               System.out.println(row.toString());
               cell1 = row.getCell(1);                //teacher's name
               cell2 = row.getCell(6);                //e-mail

               //trim()

               //login_name
               bw.write(cell2.toString().trim());
               bw.write("@mail.klcg.gov.tw");
               bw.write(",");
               //userpassword
               bw.write("password");
               bw.write(",");
               //name
               bw.write(cell1.toString());
               bw.write(",");
               //depid
               bw.write("173626");         //記得更換學校代碼
               bw.write(",");
               //email
               bw.write(cell2.toString().trim());
               bw.write("@mail.klcg.gov.tw");
               bw.write(",");
               bw.write("\r\n" );
           }
           bw.close();
           osw.close();

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
