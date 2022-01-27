package TellerTotaalOntkleeft;

import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class XLSX {
    private final XSSFWorkbook XLSXWorkbookObject;
    String filePath;
    String ontkleefDate;
    List<String> types = new ArrayList<>();
    double[] amounts = new double[0];
    List<String> statuses = new ArrayList<>();
    int typelocationInSheet = 1;
    int amountlocationInSheet = 23;

    public XLSX(String filename) throws IOException{
        filePath = "." + File.separator + filename;
        File file = new File(filePath);
        FileInputStream inputFile = new FileInputStream(file);
        this.XLSXWorkbookObject = new XSSFWorkbook(inputFile);
        inputFile.close();
    }

    public void countTotalOfEach() {
        int sheetNumber = -1;
        Scanner scanner = new Scanner(System.in);
        boolean keepLooping = true;

        while(keepLooping) {
            System.out.print("\nWelk werkblad wilt u tellen(geef de nummer van het blad): ");
            sheetNumber = scanner.nextInt();
            scanner.nextLine();
            if(sheetNumber <= 0 || sheetNumber > XLSXWorkbookObject.getNumberOfSheets()){
                System.out.println("Ongeldige pagina nummer, gelieve binnen de grenzen te blijven!");
            }
            else{
                keepLooping = false;
            }
        }
        XSSFSheet XLSXWorkSheet = XLSXWorkbookObject.getSheetAt(--sheetNumber);
        ontkleefDate = XLSXWorkSheet.getSheetName();
        int rowNumber = 0;
        //iterating over excel file
        for (Row row : XLSXWorkSheet) {
            boolean exists = false;
            int numberOfType = 0;
            if (!row.getZeroHeight()) {
                if (rowNumber != 0) {
                    String type = row.getCell(typelocationInSheet).getStringCellValue();
                    for(int i = 0; i < types.size(); i++) {
                        if(type.equals(types.get(i))){
                            exists = true;
                            numberOfType = i;
                        }
                    }
                    double amount = row.getCell(amountlocationInSheet).getNumericCellValue();
                    if(exists){
                        amounts[numberOfType] += amount;
                    }
                    else{
                        double[] tempAmount = new double[amounts.length+1];
                        System.arraycopy(amounts, 0, tempAmount, 0, amounts.length);
                        tempAmount[tempAmount.length-1] = amount;
                        amounts = tempAmount;

                        //String[] tempType = new String[types.size()+1];
                        //System.arraycopy(types, 0, tempType, 0, types.size());
                        //tempType[tempType.length-1] = type;
                        //types = tempType;

                        types.add(type);
                    }
                    Cell status = row.getCell(4);
                    String statusValue;
                    if(status == null){
                        statusValue = "blanco";
                    }
                    else{
                        statusValue = status.getStringCellValue();
                    }
                    for (String s : statuses) {
                        if (statusValue.equals(s)) {
                            exists = true;
                            break;
                        }
                    }
                    if(!exists){
                        statuses.add(statusValue);
                    }
                }
                System.out.printf("\n%d van de %d rijen geteld.", rowNumber, XLSXWorkSheet.getLastRowNum());
            }
            else{
                System.out.printf("\n%d van de %d rijen is een verborgen rij en dus niet geteld.", rowNumber, XLSXWorkSheet.getLastRowNum());
            }
            rowNumber++;
        }
        System.out.println("\n");
    }

    public void writeFile() throws IOException{
        Scanner scanner = new Scanner(System.in);
        filePath = filePath.substring(0, filePath.length()-5);
        filePath = filePath + " totalen ";
        System.out.print("Geef een achtervoegsel om toe te voegen aan de bestandsnaam (zonder \".xlsx\" er achter): ");
        String newFilenamePart = scanner.nextLine();
        filePath += newFilenamePart + ".xlsx";
        File file = new File(filePath);
        boolean test = false;
        while(!test) {
            if (!file.exists()) {
                if(file.createNewFile()) {
                    test = true;
                }
                else {
                    System.err.println("Error while creating file!");
                }
            }
            else{
                System.out.println("Naam voor nieuwe totaal file bestaat al!");
                System.out.print("Geef een achtervoegsel zoals bijvoorbeeld een datum om toe te voegen aan de bestandsnaam (zonder \".xlsx\" er achter): ");
                newFilenamePart = scanner.nextLine();
                filePath = filePath.substring(0, filePath.length()-5);
                filePath += "na " + newFilenamePart + ".xlsx";
                file = new File(filePath);
            }
        }
        XSSFWorkbook XLSXWorkbookTotals = new XSSFWorkbook();
        XSSFSheet totalSheet = XLSXWorkbookTotals.createSheet(ontkleefDate);
        totalSheet.createRow(0);
        totalSheet.getRow(0).createCell(0).setCellValue(filePath.substring(filePath.lastIndexOf(File.separator)+1, filePath.length()-5));
        totalSheet.getRow(0).createCell(1).setCellValue("");
        int sizeColumnOne = totalSheet.getRow(0).getCell(0).getStringCellValue().length();
        for (String type:types) {
            if(type.length() > sizeColumnOne){
                sizeColumnOne = type.length();
            }
        }
        sizeColumnOne += 6;

        totalSheet.addMergedRegion(new CellRangeAddress(0,0,0,1));
        totalSheet.setColumnWidth(0, 256*sizeColumnOne);
        totalSheet.autoSizeColumn(1);

        XSSFCellStyle style = XLSXWorkbookTotals.createCellStyle();
        XSSFFont font = XLSXWorkbookTotals.createFont();
        font.setFontHeight(14);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        totalSheet.getRow(0).getCell(0).setCellStyle(style);

        totalSheet.createRow(2);
        totalSheet.getRow(2).createCell(0).setCellValue("Product");
        totalSheet.getRow(2).createCell(1).setCellValue("Aantal");
        setStyle(true, totalSheet, 2, XLSXWorkbookTotals);
        for (int i = 0; i < types.size(); i++){
            totalSheet.createRow(totalSheet.getLastRowNum()+1);
            totalSheet.getRow(totalSheet.getLastRowNum()).createCell(0).setCellValue(types.get(i));
            totalSheet.getRow(totalSheet.getLastRowNum()).createCell(1).setCellValue(amounts[i]);
            setStyle(false, totalSheet, totalSheet.getLastRowNum(), XLSXWorkbookTotals);
        }
        int rowNumber = totalSheet.getLastRowNum()+4;
        totalSheet.createRow(rowNumber);
        totalSheet.getRow(rowNumber).createCell(0).setCellValue("Getelde statussen");
        setStyle(true, totalSheet, rowNumber, XLSXWorkbookTotals);
        for (String status : statuses) {
            totalSheet.createRow(totalSheet.getLastRowNum() + 1);
            totalSheet.getRow(totalSheet.getLastRowNum()).createCell(0).setCellValue(status);
            setStyle(false, totalSheet, totalSheet.getLastRowNum(), XLSXWorkbookTotals);
        }

        FileOutputStream outputStream = new FileOutputStream(file);
        XLSXWorkbookTotals.write(outputStream);
        outputStream.close();
        XLSXWorkbookObject.close();
        XLSXWorkbookTotals.close();
    }

    private void setStyle(boolean bold, XSSFSheet totalSheet, int rowNumber, XSSFWorkbook XLSXWorkbookTotals){
        for(int i = 0; i < 2; i++){
            XSSFCellStyle styleOld = totalSheet.getRow(rowNumber).getCell(i).getCellStyle();
            XSSFCellStyle newStyle = XLSXWorkbookTotals.createCellStyle();
            newStyle.cloneStyleFrom(styleOld);
            newStyle.setBorderTop(BorderStyle.THIN);
            newStyle.setBorderBottom(BorderStyle.THIN);
            newStyle.setBorderLeft(BorderStyle.THIN);
            newStyle.setBorderRight(BorderStyle.THIN);
            Font oldFont = totalSheet.getRow(rowNumber).getCell(i).getCellStyle().getFont();

            XSSFFont newFont = XLSXWorkbookTotals.createFont();
            newFont.setBold(bold);
            newFont.setFontName(newStyle.getFont().getFontName());
            newFont.setFamily(newStyle.getFont().getFamily());
            newStyle.setFont(newFont);

            totalSheet.getRow(rowNumber).getCell(i).setCellStyle(newStyle);
        }
    }
}
