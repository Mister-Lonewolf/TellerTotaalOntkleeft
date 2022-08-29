package TellerTotaalOntkleeft;

import java.io.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.util.*;


public class XLSX {
    private final XSSFWorkbook XLSXWorkbookObject;
    String filePath;
    String unstickDate;

    Map<String, Map<String, Integer>> amountPerDateAndType = new HashMap<>(); // map<typeOfSticker, map<date, amount>>
    List<String> allDates = new ArrayList<>();
    int vehicleTypeLocationInSheet = 0;
    int typeLocationInSheet = 1;
    int dateLocation = 4;
    int amountLocationInSheet = 23;

    int sheetNumber = -1;

    public XLSX(String filename) throws IOException{
        filePath = "." + File.separator + filename;
        File file = new File(filePath);
        FileInputStream inputFile = new FileInputStream(file);
        this.XLSXWorkbookObject = new XSSFWorkbook(inputFile);
        inputFile.close();
    }

    public void selectSheet(){
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
    }

    public void countPerDateAndVehicleType() {
        XSSFSheet XLSXWorkSheet = XLSXWorkbookObject.getSheetAt(--sheetNumber);
        unstickDate = XLSXWorkSheet.getSheetName();
        int rowNumber = 0;
        //iterating over excel file
        for (Row row : XLSXWorkSheet) {
            if (!row.getZeroHeight()) {
                if (rowNumber != 0 && row.getCell(0).getNumericCellValue() > 0 && row.getCell(dateLocation) != null) {
                    String type = row.getCell(typeLocationInSheet).getStringCellValue();
                    int vehicleType = (int) row.getCell(vehicleTypeLocationInSheet).getNumericCellValue();
                    Cell dateCell = row.getCell(dateLocation);
                    String dateValue = String.valueOf(CellUtils.getCellValue(dateCell));
                    int amount = (int)row.getCell(amountLocationInSheet).getNumericCellValue();
                    if((VehicleType.BUS.beginNumber <= vehicleType && VehicleType.BUS.endNumber >= vehicleType)||
                            (VehicleType.POLDER.beginNumber <= vehicleType && VehicleType.POLDER.endNumber >= vehicleType)){
                        dateValue += " BUS";
                    } else if (VehicleType.TRAM.beginNumber <= vehicleType && VehicleType.TRAM.endNumber >= vehicleType) {
                        dateValue += " TRAM";
                    }
                    if (amountPerDateAndType.containsKey(type)) {
                        Map<String, Integer> tempPerDate = amountPerDateAndType.get(type);
                        if(tempPerDate.containsKey(dateValue)) {
                            int amountTemp = amount + tempPerDate.get(dateValue);
                            tempPerDate.replace(dateValue, amountTemp);
                        }
                        tempPerDate.putIfAbsent(dateValue, amount);
                        if(!allDates.contains(dateValue)){
                            allDates.add(dateValue);
                        }
                    }
                    else {
                        if(!allDates.contains(dateValue)){
                            allDates.add(dateValue);
                        }
                        Map<String, Integer> amountPerDate = new HashMap<>();
                        amountPerDate.put(dateValue, amount);
                        amountPerDateAndType.putIfAbsent(type, amountPerDate);
                    }
                }
                System.out.printf("\n%d van de %d rijen geteld.", rowNumber, XLSXWorkSheet.getLastRowNum());
            } else {
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

        XSSFWorkbook XLSXWorkbookTotals = CellUtils.createResultWorkBook(filePath, unstickDate, amountPerDateAndType, allDates);

        FileOutputStream outputStream = new FileOutputStream(file);
        XLSXWorkbookTotals.write(outputStream);
        outputStream.close();
        XLSXWorkbookObject.close();
        XLSXWorkbookTotals.close();
    }
}
