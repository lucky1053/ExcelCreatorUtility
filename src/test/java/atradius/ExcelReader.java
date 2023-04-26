package atradius;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

public class ExcelReader {
    public File file;
    public FileInputStream fis;
    XSSFWorkbook wb;
    Sheet sheet;
    int rowCount;
    public ExcelReader() throws IOException {
        file = new File(System.getProperty("user.dir")+"\\InputFiles\\inputfile.xlsx");
        fis= new FileInputStream(file);
         wb = new XSSFWorkbook(fis);
         sheet =wb.getSheetAt(0);

    }

    public int getJobRowCount() {
         rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();
        return rowCount;
    }

    public ArrayList getJobList() {
        ArrayList list= new ArrayList();
        int counter=getJobRowCount();
        for (int i = 1; i < counter+1; i++) {
            Row row = sheet.getRow(i);
            //Create a loop to print cell values in a row
            for (int j = 0; j < row.getLastCellNum(); j++) {
                //Print Excel data in console
                list.add(row.getCell(j).getStringCellValue());
                System.out.println(row.getCell(j).getStringCellValue());
            }
        }
        return list;
}



}

