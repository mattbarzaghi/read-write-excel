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
        List<Simulate> elementiPortafoglio = readExcelFile();

        // Step 2: Convert Java Objects to JSON String
        String jsonString = convertObjects2JsonString(elementiPortafoglio);

        System.out.println(jsonString);
    }

    private static List<Simulate> readExcelFile(){
        try {
            FileInputStream excelFile = new FileInputStream("C:\\Users\\matte\\IdeaProjects\\read-write-excel\\data\\algoritmo_confronto.xlsx");
            Workbook workbook = new XSSFWorkbook(excelFile);

            Sheet sheet = workbook.getSheet("Portafoglio");
            Iterator<Row> rows = sheet.iterator();

            List<Simulate> lstPortfolio = new ArrayList<>();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if(rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();

                Simulate elementoPortafoglio = new Simulate();

                int cellIndex = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    if(cellIndex==1) {
                        elementoPortafoglio.setCodiceTipoProdottoServizio(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==2) {
                        elementoPortafoglio.setCodiceSurrogatoProdottoMIFID(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==3) {
                        elementoPortafoglio.setCodiceISIN(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==4) {
                        elementoPortafoglio.setCodiceProdotto(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==5) {
                        elementoPortafoglio.setCodiceContrattoRapporto(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==6) {
                        elementoPortafoglio.setCodiceSurrogatoProdottoServizio(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==7) {
                        elementoPortafoglio.setDescrizioneCommercialeProdotto(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==8) {
                        elementoPortafoglio.setCodiceSurrogatoProdottoServizioPadre(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==9) {
                        elementoPortafoglio.setDescrizioneCommercialeProdottoPadre(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==10) {
                        elementoPortafoglio.setCodiceTipoVersamento(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==11) {
                        elementoPortafoglio.setCodiceSostenibilitaGREEN(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==12) {
                        elementoPortafoglio.setFlagSostenibilitaECOLABEL(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==13) {
                        elementoPortafoglio.setFlagSostenibilitaPAI(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==14) {
                        elementoPortafoglio.setValoreScoreESGComplessivo(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==15) {
                        elementoPortafoglio.setValoreScoreESGComplessivo(currentCell.getStringCellValue());
                    }
                    else if(cellIndex==16) {
                        elementoPortafoglio.setDivisa(String.valueOf(currentCell.getNumericCellValue()));
                    }
                    else if(cellIndex==17) {
                        elementoPortafoglio.setDivisa(String.valueOf(currentCell.getNumericCellValue()));
                    }

                    cellIndex++;

                }

                lstPortfolio.add(elementoPortafoglio);
            }

            // Close WorkBook
            workbook.close();

            return lstPortfolio;
        } catch (IOException e) {
            throw new RuntimeException("FAIL! -> message = " + e.getMessage());
        }
    }

    private static String convertObjects2JsonString(List<Simulate> elementiPortafoglio) {
        ObjectMapper mapper = new ObjectMapper();
        String jsonString = "";

        try {
            jsonString = mapper.writeValueAsString(elementiPortafoglio);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }

        return jsonString;
    }
}
