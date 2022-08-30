package TellerTotaalOntkleeft;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
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

    public static void setStyle(boolean bold, XSSFSheet totalSheet, int rowNumber, XSSFWorkbook XLSXWorkbookTotals, int amountOfDates, boolean skipIndex, ArrayList<Integer> totals, int dateTotalSize, Boolean totalDayRow) {
        int start = skipIndex ? 1 : 0;
        int end = 1 + amountOfDates;
//        int end = skipIndexAndTotal ? amountOfDates : 1 + amountOfDates;
        for (int i = start; i <= end; i++) {
            if (totalSheet.getRow(rowNumber).getCell(i) == null) {
                totalSheet.getRow(rowNumber).createCell(i).setCellValue(0);
            }
            //XSSFCellStyle styleOld = totalSheet.getRow(rowNumber).getCell(i).getCellStyle();
            XSSFCellStyle newStyle = XLSXWorkbookTotals.createCellStyle();
            //newStyle.cloneStyleFrom(styleOld);
            newStyle.setBorderTop(BorderStyle.THIN);
            newStyle.setBorderBottom(BorderStyle.THIN);
            newStyle.setBorderLeft(BorderStyle.THIN);
            newStyle.setBorderRight(BorderStyle.THIN);
            if(i!=0) {
                newStyle.setAlignment(HorizontalAlignment.CENTER);
            }
            if(rowNumber == 2){
                newStyle.setWrapText(true);
            }
            //Font oldFont = totalSheet.getRow(rowNumber).getCell(i).getCellStyle().getFont();
            if (totals.contains(i)) {
                newStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            if (i == amountOfDates + 1 && !totalDayRow) {
                newStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
                newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            if (i == amountOfDates + 1 && totalDayRow) {
                newStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
                newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            XSSFFont newFont = XLSXWorkbookTotals.createFont();
            newFont.setBold(bold);
            newFont.setFontName(newStyle.getFont().getFontName());
            newFont.setFamily(newStyle.getFont().getFamily());
            newStyle.setFont(newFont);

            if(totals.contains(i) && rowNumber == 2){
                XSSFFont yellowFont = XLSXWorkbookTotals.createFont();
                yellowFont.setBold(bold);
                yellowFont.setFontName(newStyle.getFont().getFontName());
                yellowFont.setFamily(newStyle.getFont().getFamily());
                yellowFont.setColor(IndexedColors.YELLOW.index);
                XSSFRichTextString richString = new XSSFRichTextString(totalSheet.getRow(rowNumber).getCell(i).getStringCellValue());
                richString.applyFont(richString.length()-2, richString.length(), yellowFont);
                totalSheet.getRow(rowNumber).getCell(i).setCellValue(richString);
            }

            totalSheet.getRow(rowNumber).getCell(i).setCellStyle(newStyle);
            if (i == amountOfDates + 1) {
                totalSheet.autoSizeColumn(i);
            }
            if (i != 0 && i != amountOfDates + 1) {
                totalSheet.setColumnWidth(i, dateTotalSize);
            }
        }
    }

    public static XSSFWorkbook createResultWorkBook(String filePath, String unstickDate, Map<String, Map<String, Integer>> amountPerDateAndType, List<String> allDates, int dateTotalSize) {
        XSSFWorkbook XLSXWorkbookTotals = new XSSFWorkbook();
        XSSFSheet totalSheet = XLSXWorkbookTotals.createSheet(unstickDate);
        totalSheet.createRow(0);
        totalSheet.getRow(0).createCell(0).setCellValue(filePath.substring(filePath.lastIndexOf(File.separator) + 1, filePath.length() - 5));
        totalSheet.getRow(0).createCell(1).setCellValue("");
        int sizeColumnOne = totalSheet.getRow(0).getCell(0).getStringCellValue().length();
        for (var entry : amountPerDateAndType.entrySet()) {
            if (entry.getKey().length() > sizeColumnOne) {
                sizeColumnOne = entry.getKey().length();
            }
        }
        sizeColumnOne += 1;

        totalSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        totalSheet.setColumnWidth(0, 256 * sizeColumnOne);
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
        ArrayList<Integer> totals = new ArrayList<>();
        for (int i = 0; i <= columnNumber; i++) {
            String date = totalSheet.getRow(2).getCell(i).getStringCellValue();
            if (i != 0 && !date.contains("BUS") && !date.contains("TRAM") && !date.contains("POLDER") && !date.contains("Totaal")) {
                totals.add(i);
            }
        }
        CellUtils.setStyle(true, totalSheet, 2, XLSXWorkbookTotals, allDates.size(), false, totals, dateTotalSize, false);
        int lastRow = totalSheet.getLastRowNum();
        int totalDayRow = lastRow + 1 + amountPerDateAndType.keySet().size();
        totalSheet.createRow(totalDayRow);
//        //for (int i = 0; i < types.size(); i++){
        for (var entry : amountPerDateAndType.entrySet()) {
            totalSheet.createRow(++lastRow);
            int total = 0;
            totalSheet.getRow(lastRow).createCell(0).setCellValue(entry.getKey());
            for (Map.Entry<String, Integer> dateIntegerEntry : entry.getValue().entrySet()) {
                int locationOfDate = allDates.indexOf(dateIntegerEntry.getKey()) + 1;
                totalSheet.getRow(lastRow).createCell(locationOfDate).setCellValue(dateIntegerEntry.getValue());
                if (totalSheet.getRow(totalDayRow).getCell(locationOfDate) != null) {
                    double totalPerDay = totalSheet.getRow(totalDayRow).getCell(locationOfDate).getNumericCellValue();
                    totalPerDay += dateIntegerEntry.getValue();
                    totalSheet.getRow(totalDayRow).getCell(locationOfDate).setCellValue(totalPerDay);
                } else {
                    totalSheet.getRow(totalDayRow).createCell(locationOfDate).setCellValue(dateIntegerEntry.getValue());
                }
                if (dateIntegerEntry.getKey().contains("BUS") || dateIntegerEntry.getKey().contains("TRAM") || dateIntegerEntry.getKey().contains("POLDER")) {
                    total += dateIntegerEntry.getValue();
                }
            }
            totalSheet.getRow(lastRow).createCell(columnNumber).setCellValue(total);
            CellUtils.setStyle(false, totalSheet, lastRow, XLSXWorkbookTotals, allDates.size(), false, totals, dateTotalSize, false);
        }
        int lastTotalDayRowCell = totalSheet.getRow(totalDayRow).getLastCellNum();
        totalSheet.getRow(totalDayRow).createCell(lastTotalDayRowCell);
        int totalTotal = 0;
        for(int i = 1; i <= lastTotalDayRowCell; i++){
            totalTotal += totalSheet.getRow(totalDayRow).getCell(i).getNumericCellValue();
        }
        totalSheet.getRow(totalDayRow).getCell(lastTotalDayRowCell).setCellValue(totalTotal/2.0);
        CellUtils.setStyle(false, totalSheet, totalDayRow, XLSXWorkbookTotals, allDates.size(), true, totals, dateTotalSize, true);
        return XLSXWorkbookTotals;
    }

    public static Map<String, Integer> sortByKey(Map<String, Integer> map) {
        // Create a list from elements of HashMap
        List<Map.Entry<String, Integer>> list = new LinkedList<>(map.entrySet());

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
