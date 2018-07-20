  /*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excel.okuma;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
/**
 *
 * @author yapra_000
 */
public class excel {
    
    public static void main(String[]args)throws Exception{
    
   try {

     

    FileInputStream file = new FileInputStream(new File("C:\\test.xls"));

     

    HSSFWorkbook workbook = new HSSFWorkbook(file);


    HSSFSheet sheet = workbook.getSheetAt(0);

    

    Iterator<Row> rowIterator = sheet.iterator();
    
    while(rowIterator.hasNext()) {

        Row row = rowIterator.next();

         

        Iterator<Cell> cellIterator = row.cellIterator();

        while(cellIterator.hasNext()) {

             

            Cell cell = cellIterator.next();

             

            switch(cell.getCellType()) {

                case Cell.CELL_TYPE_BOOLEAN:

                    System.out.print(cell.getBooleanCellValue() + "\t\t");

                    break;

                case Cell.CELL_TYPE_NUMERIC:

                    System.out.print(cell.getNumericCellValue() + "\t\t");

                    break;

                case Cell.CELL_TYPE_STRING:

                    System.out.print(cell.getStringCellValue() + "\t\t");

                    break;

            }

        }

        System.out.println("");

    }

    file.close();

    FileOutputStream out = 

        new FileOutputStream(new File("C:\\test.xls"));

    workbook.write(out);

    out.close();

     

} catch (FileNotFoundException e) {

    e.printStackTrace();

} catch (IOException e) {

    e.printStackTrace();

}
    
    
    }
    
}
