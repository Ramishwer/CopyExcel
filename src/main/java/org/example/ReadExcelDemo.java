package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.text.DateFormat;
import java.text.Format;
import java.text.SimpleDateFormat;


//import statements
public class ReadExcelDemo
{
    public static void main(String[] args)
    {
        try
        {
            FileInputStream file = new FileInputStream(new File("D:\\New folder\\New folder\\Book1.xlsx"));

            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);

            String target="Name";
            String target1="10/28/01";

            for(int i=0;i<=sheet.getLastRowNum()+1;i++)
            {
                Row row = sheet.getRow(i);


                for(int j=0;j<sheet.getLastRowNum();j++)
                {
                    Cell cell = row.getCell(j);

                    DataFormatter dataFormatter = new DataFormatter();
                    String value = dataFormatter.formatCellValue(cell);


                    if(value.equals(target) || value.equals(target1)){
                     System.out.println(cell.getRowIndex()+" "+cell.getColumnIndex());
                    }

                    if(cell.getRowIndex()==3 && cell.getColumnIndex()==1){
                        System.out.println(cell.getStringCellValue());
                    }


//                    switch (cell.getCellType())
//                    {
//                        case Cell.CELL_TYPE_NUMERIC:
//                 //   System.out.print(cell. + " ");
//                            break;
//                        case Cell.CELL_TYPE_STRING:
//                    //    System.out.print(cell.getStringCellValue() + " ");
//                            break;
//                        case Cell.CELL_TYPE_BLANK:
//                 //        System.out.println(cell.getStringCellValue()+" ");
//                            break;
//                        case Cell.CELL_TYPE_BOOLEAN:
//                            //            System.out.println(cell.getBooleanCellValue()+" ");
//                            break;
//                        case Cell.CELL_TYPE_ERROR:
//                            System.out.println(cell.getErrorCellValue()+" ");
//                            break;
//                        case Cell.CELL_TYPE_FORMULA:
//                            System.out.println(cell.getCellFormula()+" ");
//                            break;
//
//                    }


                }
                System.out.println("");

            }
            file.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

//    static int[] getindex(Cell cell,String target){
//     int cellcount=ro
//
//        for(int i=0;i<cell;i++){
//            for(int j=0;j<cellcount;j++){
//                if(row.equals(target)){
//                    return new int[] { i, j };
//                }
//            }
//        }
//        return new int[] { -1, -1 };
//    }


}