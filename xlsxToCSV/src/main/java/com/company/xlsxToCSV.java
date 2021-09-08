package com.company;

import java.io.File;
import java.io.FileDescriptor;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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

           Workbook workbook = null;

           String ext  = FilenameUtils.getExtension(inputFile.toString());

           if(ext.equalsIgnoreCase("xlsx")) {
               workbook = new XSSFWorkbook(fis);
           }else if(ext.equalsIgnoreCase("xls")){
               workbook = new HSSFWorkbook(fis);
           }





       } catch(Exception e) {
           System.out.println("*****read excel*****");
           e.printStackTrace();
       }
   }



    static void xlsx(File inputFile, File outputFile) {
        StringBuffer data = new StringBuffer();


        try {
            FileOutputStream fos = new FileOutputStream(outputFile);
            FileInputStream fis = new FileInputStream(inputFile);

            Workbook workbook = null;

            String ext  = FilenameUtils.getExtension(inputFile.toString());

            if(ext.equalsIgnoreCase("xlsx")) {
                workbook = new XSSFWorkbook(fis);
            }else if(ext.equalsIgnoreCase("xls")){
                workbook = new HSSFWorkbook(fis);
            }
            // get first sheet from the workbook

            int numberOfSheets = workbook.getNumberOfSheets();
            Row row;
            Cell cell;


            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(0);


                Iterator<Row> rowIterator = sheet.iterator();

                while (rowIterator.hasNext()) {
                    row = rowIterator.next();

                    // For each row, iterate through each columns
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {

                        cell = cellIterator.next();


//                        String str;
//                        str = (String) cellIterator.next().toString().subSequence(0,1);
//                        System.out.println(str);



                        switch (cell.getCellType()) {
                            case BOOLEAN:
                                data.append(cell.getBooleanCellValue() + ",");

                                break;
                            case NUMERIC:
                                data.append(cell.getNumericCellValue() + ",");

                                break;
                            case STRING:
                                data.append(cell.getStringCellValue() + ",");
//                                System.out.println(cell.getStringCellValue());
                                break;

                            case BLANK:
                                data.append("" + ",");
                                break;
                            default:
                                data.append(cell + ",");

                        }
                    }
                    data.append('\n'); // appending new line after each row
                }
            }
            fos.write(data.toString().getBytes());
            fos.close();

        } catch(Exception ioe) {
            ioe.printStackTrace();
        }

    }

    public static void main(String[] args) {
        // reading file from desktop
        File inputFile = new File("/home/user/Downloads/govMail/七堵國小.xlsx"); //provide your path
        // writing excel data to csv
        File outputFile = new File("/home/user/Downloads/govMail/七堵國小.csv");  //provide your path
        //work space
        xlsx(inputFile, outputFile);

        System.out.println("Conversion of " + inputFile + " to flat file: "
                + outputFile + " is completed");
    }

}
