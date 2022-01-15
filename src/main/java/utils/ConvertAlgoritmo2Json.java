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

public class ConvertAlgoritmo2Json {

    public static void main(String[] args) {
        // Step 1: Read Excel File into Java List Objects
        List<BmedProduct> bmedProductElel = readExcelFile();

        // Step 2: Convert Java Objects to JSON String
        String jsonString = convertObjects2JsonString(bmedProductElel);

        System.out.println(jsonString);
    }

    private static List<BmedProduct> readExcelFile(){
        try {
            FileInputStream excelFile = new FileInputStream("C:\\Users\\matte\\IdeaProjects\\read-write-excel\\data\\algoritmo.xlsx");
            Workbook workbook = new XSSFWorkbook(excelFile);

            Sheet sheet = workbook.getSheet("Portafoglio");
            Iterator<Row> rows = sheet.iterator();

            List<BmedProduct> bmedProductList = new ArrayList<>();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if(rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();

                BmedProduct bmedProduct = new BmedProduct();

                int cellIndex = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    if(cellIndex==1) {
                        bmedProduct.setCodiceTipoProdottoServizio(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==2) {
                        bmedProduct.setCodiceSurrogatoProdottoMIFID(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==3) {
                        bmedProduct.setCodiceISIN(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==4) {
                        bmedProduct.setCodiceProdotto(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==5) {
                        bmedProduct.setCodiceContrattoRapporto(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==6) {
                        bmedProduct.setCodiceSurrogatoProdottoServizio(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==7) {
                        bmedProduct.setDescrizioneCommercialeProdotto(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==8) {
                        bmedProduct.setCodiceSurrogatoProdottoServizioPadre(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==9) {
                        bmedProduct.setDescrizioneCommercialeProdottoPadre(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==10) {
                        bmedProduct.setCodiceTipoVersamento(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==11) {
                        bmedProduct.setCodiceSostenibilitaGREEN(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==12) {
                        bmedProduct.setFlagSostenibilitaECOLABEL(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==13) {
                        bmedProduct.setFlagSostenibilitaPAI(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==14) {
                        bmedProduct.setValoreScoreESGComplessivo(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==15) {
                        bmedProduct.setValoreScoreESGComplessivo(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==16) {
                        bmedProduct.setDivisa(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==17) {
                        bmedProduct.setDivisa(String.valueOf(currentCell.getNumericCellValue()));
                    }

                    cellIndex++;

                }

                bmedProductList.add(bmedProduct);
            }

            // Close WorkBook
            workbook.close();

            return bmedProductList;
        } catch (IOException e) {
            throw new RuntimeException("FAIL! -> message = " + e.getMessage());
        }
    }

    private static String convertObjects2JsonString(List<BmedProduct> bmedProductList) {
        ObjectMapper mapper = new ObjectMapper();
        String jsonString = "";

        try {
            jsonString = mapper.writeValueAsString(bmedProductList);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }

        return jsonString;
    }
}
