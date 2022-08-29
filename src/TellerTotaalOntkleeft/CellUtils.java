package TellerTotaalOntkleeft;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

import java.io.File;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

public class CellUtils {
    public static Object getCellValue(Cell cell) {
        Objects.requireNonNull(cell, "cell is null");

        CellType cellType = cell.getCellTypeEnum();
        if (cellType == CellType.BLANK) {
            return "blanco";
        } else if (cellType == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        } else if (cellType == CellType.ERROR) {
            throw new RuntimeException("Error cell is unsupported");
        } else if (cellType == CellType.FORMULA) {
            throw new RuntimeException("Formula cell is unsupported");
        } else if (cellType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                Date date = cell.getDateCellValue();
                SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
                return formatter.format(date);
            } else {
                return cell.getNumericCellValue();
            }
        } else if (cellType == CellType.STRING) {
            return cell.getStringCellValue();
        } else {
            throw new RuntimeException("Unknow type cell");
        }
    }

    static void colorOrangeOrYellow(XSSFWorkbook XLSXWorkbookObject, XSSFCellStyle style, int cellNumber, boolean colorNumber) {
        if (colorNumber) {
            style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
            if (cellNumber != 0) {
                XSSFFont newFont = XLSXWorkbookObject.createFont();
                newFont.setBold(true);
                newFont.setFontName(style.getFont().getFontName());
                newFont.setFontHeightInPoints(style.getFont().getFontHeightInPoints());
                newFont.setFamily(style.getFont().getFamily());
                style.setFont(newFont);
            }
        } else {
            style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        }
    }

    public static void setStyle(boolean bold, XSSFSheet totalSheet, int rowNumber, XSSFWorkbook XLSXWorkbookTotals, int amountOfDates, boolean skipIndexAndTotal){
        int start = skipIndexAndTotal?1:0;
        int end = skipIndexAndTotal?amountOfDates:1+amountOfDates;
        ArrayList<Integer> totals = new ArrayList<>();
        for(int i = start; i <= end; i++) {
            if (totalSheet.getRow(rowNumber).getCell(i) == null) {
                totalSheet.getRow(rowNumber).createCell(i).setCellValue(0);
            }
            if(rowNumber == 2){
                String date = totalSheet.getRow(rowNumber).getCell(i).getStringCellValue();
                if(i!=0 && !date.contains("BUS") && !date.contains("TRAM") && !date.contains("POLDER")){
                    totals.add(i);
                }
            }
            //XSSFCellStyle styleOld = totalSheet.getRow(rowNumber).getCell(i).getCellStyle();
            XSSFCellStyle newStyle = XLSXWorkbookTotals.createCellStyle();
            //newStyle.cloneStyleFrom(styleOld);
            newStyle.setBorderTop(BorderStyle.THIN);
            newStyle.setBorderBottom(BorderStyle.THIN);
            newStyle.setBorderLeft(BorderStyle.THIN);
            newStyle.setBorderRight(BorderStyle.THIN);
            //Font oldFont = totalSheet.getRow(rowNumber).getCell(i).getCellStyle().getFont();
            if(totals.contains(i)){
                XSSFColor orange = new XSSFColor();
                orange.setRGB(new byte[] {(byte) 255, (byte) 165, (byte) 0, (byte) 255});
                newStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, orange);
            }

            XSSFFont newFont = XLSXWorkbookTotals.createFont();
            newFont.setBold(bold);
            newFont.setFontName(newStyle.getFont().getFontName());
            newFont.setFamily(newStyle.getFont().getFamily());
            newStyle.setFont(newFont);

            totalSheet.getRow(rowNumber).getCell(i).setCellStyle(newStyle);
            if (i != 0) {
                totalSheet.autoSizeColumn(i);
            }
        }
    }

    public static XSSFWorkbook createResultWorkBook(String filePath, String unstickDate, Map<String, Map<String, Integer>> amountPerDateAndType, List<String> allDates){
        XSSFWorkbook XLSXWorkbookTotals = new XSSFWorkbook();
        XSSFSheet totalSheet = XLSXWorkbookTotals.createSheet(unstickDate);
        totalSheet.createRow(0);
        totalSheet.getRow(0).createCell(0).setCellValue(filePath.substring(filePath.lastIndexOf(File.separator)+1, filePath.length()-5));
        totalSheet.getRow(0).createCell(1).setCellValue("");
        int sizeColumnOne = totalSheet.getRow(0).getCell(0).getStringCellValue().length();
        for (var entry : amountPerDateAndType.entrySet()) {
            if(entry.getKey().length() > sizeColumnOne){
                sizeColumnOne = entry.getKey().length();
            }
        }
        sizeColumnOne += 1;

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
        int columnNumber = 0;
        allDates.sort(String::compareTo);
        totalSheet.getRow(2).createCell(columnNumber).setCellValue("Product");
        for (String date : allDates) {
            totalSheet.getRow(2).createCell(++columnNumber).setCellValue(date);
        }
        totalSheet.getRow(2).createCell(++columnNumber).setCellValue("Totaal");
        CellUtils.setStyle(true, totalSheet, 2, XLSXWorkbookTotals, allDates.size(), false);
        int lastRow = totalSheet.getLastRowNum();
        int totalDayRow = lastRow+1+amountPerDateAndType.keySet().size();
        totalSheet.createRow(totalDayRow);
//        //for (int i = 0; i < types.size(); i++){
        for (var entry : amountPerDateAndType.entrySet()) {
            totalSheet.createRow(++lastRow);
            int total = 0;
            totalSheet.getRow(lastRow).createCell(0).setCellValue(entry.getKey());
            for (Map.Entry<String, Integer> dateIntegerEntry : entry.getValue().entrySet()) {
                int locationOfDate = allDates.indexOf(dateIntegerEntry.getKey())+1;
                totalSheet.getRow(lastRow).createCell(locationOfDate).setCellValue(dateIntegerEntry.getValue());
                if(totalSheet.getRow(totalDayRow).getCell(locationOfDate) != null){
                    double totalPerDay = totalSheet.getRow(totalDayRow).getCell(locationOfDate).getNumericCellValue();
                    totalPerDay += dateIntegerEntry.getValue();
                    totalSheet.getRow(totalDayRow).getCell(locationOfDate).setCellValue(totalPerDay);
                }
                else {
                    totalSheet.getRow(totalDayRow).createCell(locationOfDate).setCellValue(dateIntegerEntry.getValue());
                }
                if(dateIntegerEntry.getKey().contains("BUS")||dateIntegerEntry.getKey().contains("TRAM")||dateIntegerEntry.getKey().contains("POLDER")) {
                    total += dateIntegerEntry.getValue();
                }
            }
            totalSheet.getRow(lastRow).createCell(columnNumber).setCellValue(total);
            CellUtils.setStyle(false, totalSheet, lastRow, XLSXWorkbookTotals, allDates.size(), false);
        }
        CellUtils.setStyle(false, totalSheet, totalDayRow, XLSXWorkbookTotals, allDates.size(), true);

//
//        totalSheet.getRow(totalSheet.getLastRowNum()+1);

//        int rowNumber = totalSheet.getLastRowNum()+4;
//        totalSheet.createRow(rowNumber);
//        totalSheet.getRow(rowNumber).createCell(0).setCellValue("Statussen");
//        totalSheet.getRow(rowNumber).createCell(1).setCellValue("Aantal");
//        CellUtils.setStyle(true, totalSheet, rowNumber, XLSXWorkbookTotals);
//        statuses = sortByKey(statuses);
//        for (Map.Entry<String, Map<Date, Integer>> stringMapEntry : amountPerDateAndType.entrySet()) {
//            totalSheet.createRow(totalSheet.getLastRowNum() + 1);
//            totalSheet.getRow(totalSheet.getLastRowNum()).createCell(0).setCellValue(entry.getKey());
//            totalSheet.getRow(totalSheet.getLastRowNum()).createCell(1).setCellValue(entry.getValue());
//            CellUtils.setStyle(false, totalSheet, totalSheet.getLastRowNum(), XLSXWorkbookTotals);
//        }
        return XLSXWorkbookTotals;
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
