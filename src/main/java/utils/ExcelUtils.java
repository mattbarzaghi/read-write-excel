package utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

    public static void main(String[] args) {
        getRowCount();
        getCellData();
    }

    // function to get row count
    public static void getRowCount(){

        String projDir = System.getProperty("user.dir");
        System.out.println(projDir);

        try {
            String excelPath = projDir + "\\data\\eisenhower.xlsx";
            System.out.println(excelPath);
            // very important to add excel path in the new statement
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(excelPath);
            XSSFSheet sheet = xssfWorkbook.getSheet("Sheet1");
            int rowCount = sheet.getPhysicalNumberOfRows();
            System.out.println("N° of rows: " + rowCount);

        } catch (Exception e){
            System.out.println(e.getMessage());
            e.printStackTrace();
        }

    }

    // function to get row count
    public static void getCellData(){

        String projDir = System.getProperty("user.dir");
        System.out.println(projDir);

        try {

            String excelPath = projDir + "\\data\\eisenhower.xlsx";
            System.out.println(excelPath);
            // very important to add excel path in the new statement
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(excelPath);
            XSSFSheet sheet = xssfWorkbook.getSheet("Sheet1");

            String cellValue = sheet.getRow(1).getCell(0).getStringCellValue();
            System.out.println(cellValue);

        } catch (Exception e){
            System.out.println(e.getMessage());
            e.printStackTrace();
        }

    }

}
