package TellerTotaalOntkleeft;

import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.util.*;

public class XLSX {
    private final XSSFWorkbook XLSXWorkbookObject;
    String filePath;
    String ontkleefDate;
    Map<String, Double> types = new HashMap<>();
    //double[] amounts = new double[0];
    Map<String, Integer> statuses = new HashMap<>();
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
            //boolean exists = false;
            //int numberOfType = 0;
            if (!row.getZeroHeight()) {
                if (rowNumber != 0) {
                    String type = row.getCell(typelocationInSheet).getStringCellValue();
                    /*for(int i = 0; i < types.size(); i++) {
                        if(type.equals(types.get(i))){
                            exists = true;
                            numberOfType = i;
                        }
                    }*/
                    double amount = row.getCell(amountlocationInSheet).getNumericCellValue();
                    if(types.containsKey(type)){//if(exists){
                        double amountTemp = amount + types.get(type);
                        types.replace(type, amountTemp);
                    }
                    types.putIfAbsent(type, amount);

                    /*else{
                        double[] tempAmount = new double[amounts.length+1];
                        System.arraycopy(amounts, 0, tempAmount, 0, amounts.length);
                        tempAmount[tempAmount.length-1] = amount;
                        amounts = tempAmount;

                        String[] tempType = new String[types.size()+1];
                        System.arraycopy(types, 0, tempType, 0, types.size());
                        tempType[tempType.length-1] = type;
                        types = tempType;
                        types.add(type);
                    }*/
                    Cell status = row.getCell(4);
                    String statusValue = String.valueOf(CellUtils.getCellValue(status));

                    /*for (String s : statuses) {
                        if (statusValue.equals(s)) {
                            exists = true;
                            break;
                        }
                    }*/
                    if(statuses.containsKey(statusValue)){//if(!exists){
                        //statuses.add(statusValue);
                        int amountOfStatuses = statuses.get(statusValue) + (int)amount;
                        statuses.put(statusValue, amountOfStatuses);
                    }
                    statuses.putIfAbsent(statusValue, (int)amount);
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
        String output = "Geef een achtervoegsel om toe te voegen aan de bestandsnaam \"" + filePath + "\" (zonder \".xlsx\" er achter): ";
        System.out.print(output);
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
                System.out.print(output);
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
        for (Map.Entry<String, Double> entry : types.entrySet()) {
            if(entry.getKey().length() > sizeColumnOne){
                sizeColumnOne = entry.getKey().length();
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
        CellUtils.setStyle(true, totalSheet, 2, XLSXWorkbookTotals);
        //for (int i = 0; i < types.size(); i++){
        for (Map.Entry<String, Double> entry : types.entrySet()) {
            totalSheet.createRow(totalSheet.getLastRowNum()+1);
            totalSheet.getRow(totalSheet.getLastRowNum()).createCell(0).setCellValue(entry.getKey());
            totalSheet.getRow(totalSheet.getLastRowNum()).createCell(1).setCellValue(entry.getValue());
            CellUtils.setStyle(false, totalSheet, totalSheet.getLastRowNum(), XLSXWorkbookTotals);
        }
        int rowNumber = totalSheet.getLastRowNum()+4;
        totalSheet.createRow(rowNumber);
        totalSheet.getRow(rowNumber).createCell(0).setCellValue("Statussen");
        totalSheet.getRow(rowNumber).createCell(1).setCellValue("Aantal");
        CellUtils.setStyle(true, totalSheet, rowNumber, XLSXWorkbookTotals);
        //for (String status : statuses) {
        statuses = sortByKey(statuses);
        for (Map.Entry<String, Integer> entry : statuses.entrySet()) {
            totalSheet.createRow(totalSheet.getLastRowNum() + 1);
            totalSheet.getRow(totalSheet.getLastRowNum()).createCell(0).setCellValue(entry.getKey());
            totalSheet.getRow(totalSheet.getLastRowNum()).createCell(1).setCellValue(entry.getValue());
            CellUtils.setStyle(false, totalSheet, totalSheet.getLastRowNum(), XLSXWorkbookTotals);
        }

        FileOutputStream outputStream = new FileOutputStream(file);
        XLSXWorkbookTotals.write(outputStream);
        outputStream.close();
        XLSXWorkbookObject.close();
        XLSXWorkbookTotals.close();
    }

    public static Map<String, Integer> sortByKey(Map<String, Integer> map){
        // Create a list from elements of HashMap
        List<Map.Entry<String, Integer> > list = new LinkedList<>(map.entrySet());

        // Sort the list using lambda expression
        list.sort(Map.Entry.comparingByKey());

        // put data from sorted list to hashmap
        HashMap<String, Integer> temp = new LinkedHashMap<>();
        for (Map.Entry<String, Integer> aa : list) {
            temp.put(aa.getKey(), aa.getValue());
        }
        return temp;
    }
}
