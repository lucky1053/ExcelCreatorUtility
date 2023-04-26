package atradius;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;

public class BaseClass {
    static LinkedHashMap<String, String> map;
    public static void main(String[] args) throws IOException {
        ExcelReader er= new ExcelReader();
        OutputExcelWriter oew= new OutputExcelWriter();

       XSSFWorkbook wb= oew.verfiyFilePresence();

        ArrayList<String> list = er.getJobList();

        int count = er.getJobRowCount();
        System.out.println(count);

        //
        GetUIDetails guidetails= new GetUIDetails();
        for(int i=0;i<count;i++){
           map= (LinkedHashMap<String, String>) guidetails.prepareTestData(list.get(i));
            System.out.println("Hashmap : "+map);
            oew.writeDatainExcel(map);
        }





    }
}
