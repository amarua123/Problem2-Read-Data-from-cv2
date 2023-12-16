package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;

public class Main {
    public static void main(String[] args) {
        try
        {
            FileInputStream file = new FileInputStream("C:\\Users\\amar.sarkar\\Downloads\\Problem2-Read-Data-from-cv2\\src\\main\\resources\\Data.xlsx");

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook wb = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet ws = wb.getSheetAt(0);
            DateFormat df1 = new SimpleDateFormat("dd-MM-yyyy");
            DateFormat df2 = new SimpleDateFormat("hh:mm:ssa");
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = ws.iterator();
            rowIterator.next(); //skipping 1st row which has headings

            while (rowIterator.hasNext())
            {

                Row row = rowIterator.next();
                //For each row, iterate through all the columns
//                Iterator<Cell> cellIterator = row.cellIterator();

                if(row.getCell(0).getCellType() == STRING){
                    continue;
                }
                if(row.getCell(6).getCellType() == STRING){
                    continue;
                }
                try{
                    String cellDate = df1.format(row.getCell(0).getDateCellValue());
                    String team = row.getCell(2).getStringCellValue();
                    String panel = row.getCell(3).getStringCellValue();
                    String round = row.getCell(4).getStringCellValue();
                    String skill = row.getCell(5).getStringCellValue();
                    String time = df2.format(row.getCell(6).getDateCellValue());
                    String candidate_cur_loc = row.getCell(7).getStringCellValue();
                    String candidate_pref_loc = row.getCell(8).getStringCellValue();
                    String candidate_name = row.getCell(9).getStringCellValue();
                    System.out.print(cellDate+" "+team+" "+panel+" "+round+" "+skill+" "+time+" ");
                    System.out.print(candidate_cur_loc+" "+candidate_pref_loc+" "+candidate_name);
                }catch (NullPointerException newe){
                    continue;
                }

                System.out.println();
//                while (cellIterator.hasNext())
//                {
//                    Cell cell = cellIterator.next();
//
//                    //Check the cell type and format accordingly
//                    switch (cell.getCellType())
//                    {
//                        case NUMERIC:
//                            System.out.print(df1.format(cell.getDateCellValue())+" ");
//                            break;
//                        case STRING:
//                            if(cell.getStringCellValue().compareTo("") != 0){
//                                System.out.print(cell.getStringCellValue()+" ");
//                            }
//                            break;
//
//                    }
//                }
//                System.out.println("Reading File Completed.");
            }
            file.close();
        }
        catch (Exception ex)
        {
            ex.printStackTrace();
        }
    }
}