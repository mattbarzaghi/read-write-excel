package utils;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ConvertExcel2Json {

    public static void main(String[] args) {
        // Step 1: Read Excel File into Java List Objects
        List<Portafoglio> portafoglios = readExcelFile();

        // Step 2: Convert Java Objects to JSON String
        String jsonString = convertObjects2JsonString(portafoglios);

        System.out.println(jsonString);
    }

    private static List<Portafoglio> readExcelFile(){
        try {
            FileInputStream excelFile = new FileInputStream("C:\\Users\\matte\\IdeaProjects\\read-write-excel\\data\\customers-1.xlsx");
            Workbook workbook = new XSSFWorkbook(excelFile);

            Sheet sheet = workbook.getSheet("Customers");
            Iterator<Row> rows = sheet.iterator();

            List<Portafoglio> lstPortafoglios = new ArrayList<>();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if(rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();

                Portafoglio cust = new Portafoglio();

                int cellIndex = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    if(cellIndex==0) { // ID
                        cust.setId(String.valueOf(currentCell.getNumericCellValue()));
                    } else if(cellIndex==1) { // Name
                        cust.setName(currentCell.getStringCellValue());
                    } else if(cellIndex==2) { // Address
                        cust.setAddress(currentCell.getStringCellValue());
                    } else if(cellIndex==3) { // Age
                        cust.setAge(String.valueOf(currentCell.getNumericCellValue()));
                    }

                    cellIndex++;
                }

                lstPortafoglios.add(cust);
            }

            // Close WorkBook
            workbook.close();

            return lstPortafoglios;
        } catch (IOException e) {
            throw new RuntimeException("FAIL! -> message = " + e.getMessage());
        }
    }

    private static String convertObjects2JsonString(List<Portafoglio> portafoglios) {
        ObjectMapper mapper = new ObjectMapper();
        String jsonString = "";

        try {
            jsonString = mapper.writeValueAsString(portafoglios);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }

        return jsonString;
    }
}