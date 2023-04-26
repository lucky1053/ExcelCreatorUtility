package atradius;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.LinkedHashMap;

public class OutputExcelWriter {
    String filePathString;
    public XSSFWorkbook wb;

    public OutputExcelWriter() {
        filePathString = System.getProperty("user.dir") + "\\OutputFiles\\OutputFIle.xlsx";
     }

    public void writeDatainExcel( LinkedHashMap<String, String> map) throws IOException {
        File file = new File(System.getProperty("user.dir")+"\\OutputFiles\\OutputFIle.xlsx");
        FileInputStream fis=new FileInputStream(file);
        //creating workbook instance that refers to .xls file
        XSSFWorkbook wb=new XSSFWorkbook(fis);
        String value =null;

        for ( String k : map.keySet()) {

            if (k.contains("Job")) {
                System.out.println("Key " + k + " contains the word " + "Job Name" + ".");
                value=map.get(k);
            }
        }
        System.out.println("Sheetname is : "+value);

        if (wb.getSheet(value) == null) {
         XSSFSheet sheet=   wb.createSheet(value);
            System.out.println("Sheet created with name: " + value);
            Row headerRow = sheet.createRow(0);
            String[] headers = {"Job Name", "Start Date Name", "End Date Name", "Record Passed"};
            XSSFCellStyle headerStyle = wb.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderBottom(BorderStyle.MEDIUM);
            headerStyle.setBorderTop(BorderStyle.MEDIUM);
            headerStyle.setBorderLeft(BorderStyle.MEDIUM);
            headerStyle.setBorderRight(BorderStyle.MEDIUM);
            XSSFFont headerFont = wb.createFont();
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);
            for (int j = 0; j < headers.length; j++) {
                Cell headerCell = headerRow.createCell(j);
                headerCell.setCellValue(headers[j]);
                headerCell.setCellStyle(headerStyle);
            }

            int rowindex=1;
            XSSFCellStyle dataStyle = wb.createCellStyle();
            dataStyle.setBorderBottom(BorderStyle.THIN);
            dataStyle.setBorderTop(BorderStyle.THIN);
            dataStyle.setBorderLeft(BorderStyle.THIN);
            dataStyle.setBorderRight(BorderStyle.THIN);
            Cell cell=null;
            for(String key : map.keySet()){
                String v = map.get(key);
                if(key.startsWith("Job Name")){
                    Row datarow= sheet.createRow(rowindex++);
                    cell= datarow.createCell(0);
                    cell.setCellStyle(dataStyle);
                    cell.setCellValue(v);


                    String startkey = "Start Date Name_"+key.substring(9);
                    String startvalue= map.get(startkey);
                    cell= datarow.createCell(1);
                    cell.setCellStyle(dataStyle);
                    cell.setCellValue(startvalue);
                    sheet.autoSizeColumn(1);


                    String endkey = "End Date Name_"+key.substring(9);
                    String endvalue= map.get(endkey);
                    cell= datarow.createCell(2);
                    cell.setCellStyle(dataStyle);
                    cell.setCellValue(startvalue);
                    sheet.autoSizeColumn(2);

                    String recordpassedkey = "Record Passed_"+key.substring(9);
                    String recordpassedvalue= map.get(recordpassedkey);
                    cell= datarow.createCell(3);
                    cell.setCellStyle(dataStyle);
                    cell.setCellValue(startvalue);
                    sheet.autoSizeColumn(3);


                }



            }

            FileOutputStream out = new FileOutputStream(file);
            wb.write(out);
            out.close();
            wb.close();
            System.out.println("File created successfully at " + filePathString);
        } else {
            System.out.println("Sheet already exists with name: " + value);
            XSSFSheet sheet= wb.getSheet(value);
           // Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);

            int rowindex=sheet.getLastRowNum()+1;
            XSSFCellStyle dataStyle = wb.createCellStyle();
            dataStyle.setBorderBottom(BorderStyle.THIN);
            dataStyle.setBorderTop(BorderStyle.THIN);
            dataStyle.setBorderLeft(BorderStyle.THIN);
            dataStyle.setBorderRight(BorderStyle.THIN);
            Cell cell=null;
            for(String key : map.keySet()){
                String v = map.get(key);
                if(key.startsWith("Job Name")){
                    Row datarow= sheet.createRow(rowindex++);

                    cell= datarow.createCell(0);
                    cell.setCellStyle(dataStyle);
                    cell.setCellValue(v);

                    String startkey = "Start Date Name_"+key.substring(9);
                    String startvalue= map.get(startkey);
                    cell= datarow.createCell(1);
                    cell.setCellStyle(dataStyle);
                    cell.setCellValue(startvalue);
                    sheet.autoSizeColumn(1);


                    String endkey = "End Date Name_"+key.substring(9);
                    String endvalue= map.get(endkey);
                    cell= datarow.createCell(2);
                    cell.setCellStyle(dataStyle);
                    cell.setCellValue(startvalue);
                    sheet.autoSizeColumn(2);

                    String recordpassedkey = "Record Passed_"+key.substring(9);
                    String recordpassedvalue= map.get(recordpassedkey);
                    cell= datarow.createCell(3);
                    cell.setCellStyle(dataStyle);
                    cell.setCellValue(startvalue);
                    sheet.autoSizeColumn(3);


                }



            }

            FileOutputStream out = new FileOutputStream(file);
            wb.write(out);
            out.close();
            wb.close();
            System.out.println("File created successfully at " + filePathString);
        }
    }

    public XSSFWorkbook verfiyFilePresence() {
        boolean flag = false;
        File f = new File(filePathString);
        if (f.exists() && !f.isDirectory()) {
            flag = true;
        }
        XSSFWorkbook workbook = null;
        if (!flag) {
            try {
                File file = new File(filePathString);
                workbook = new XSSFWorkbook();
                workbook.createSheet("Sheet 1");


                // Write the workbook to the file
                FileOutputStream out = new FileOutputStream(file);
                workbook.write(out);
                out.close();
                workbook.close();
                System.out.println("File created successfully at " + filePathString);
            } catch (Exception e) {
                e.printStackTrace();
            }

        }
        else{
            File file = new File(filePathString);
            workbook = new XSSFWorkbook();
        }
        return workbook;
    }
}
