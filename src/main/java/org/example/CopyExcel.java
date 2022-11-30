package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopyExcel {

    public static void main(String[] args) throws IOException
    {

        File inputFile=new File("D:\\New folder\\New folder\\demo.xlsx");
        FileInputStream fis=new FileInputStream(inputFile);

        XSSFWorkbook inputWorkbook=new XSSFWorkbook(fis);

        int inputSheetCount=inputWorkbook.getNumberOfSheets();
        System.out.println("Input sheetCount: "+inputSheetCount);

        for(int i=0;i<inputSheetCount;i++)
        {
            XSSFSheet inputSheet=inputWorkbook.getSheetAt(i);

            int outrowindex = inputSheet.getLastRowNum();

            for(int j=0;j<=outrowindex;j++) {

                XSSFRow inputrow = inputSheet.getRow(j);
                XSSFRow outputrow = inputSheet.createRow(inputrow.getRowNum()+outrowindex+1);

                copyRow(inputrow,outputrow);
            }
        }



        try {

            File yourFile = new File("D:\\New folder\\New folder\\demo.xlsx");

            if (!yourFile.exists()) {
                yourFile.createNewFile();
            }
            FileOutputStream out = new FileOutputStream(yourFile, false);

            inputWorkbook.write(out);
        }
        catch(Exception e) {
            e.printStackTrace();
        }
        fis.close();

    }

    public static void copyRow(XSSFRow inputrow,XSSFRow outputrow)
    {
        int cellcount=inputrow.getLastCellNum();

        System.out.println(cellcount+" cols in inputsheet "+inputrow.getSheet().getSheetName());

        for( int i=0;i<cellcount;i++) {


            Cell inputcell= inputrow.getCell(i);


            Cell outputcell=outputrow.createCell(i);


            switch(inputcell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    outputcell.setCellValue(inputcell.getStringCellValue());
                    System.out.println(inputcell.getStringCellValue());
                    break;

                case Cell.CELL_TYPE_BOOLEAN:
                    outputcell.setCellValue(inputcell.getBooleanCellValue());
                    System.out.println(inputcell.getBooleanCellValue());
                    break;

                case Cell.CELL_TYPE_ERROR:
                    outputcell.setCellErrorValue(inputcell.getErrorCellValue());
                    System.out.println(inputcell.getErrorCellValue());
                    break;

                case Cell.CELL_TYPE_FORMULA:
                    outputcell.setCellFormula(inputcell.getCellFormula());
                    System.out.println(inputcell.getCellFormula());
                    break;

                case Cell.CELL_TYPE_NUMERIC:
                    outputcell.setCellValue(inputcell.getNumericCellValue());
                    System.out.println(inputcell.getNumericCellValue());
                    break;

                case Cell.CELL_TYPE_STRING:
                    outputcell.setCellValue(inputcell.getStringCellValue());
                    System.out.println(inputcell.getStringCellValue());
                    break;

                default:
                    System.out.println("Default Values");
                    break;
            }

        }

    }

}

