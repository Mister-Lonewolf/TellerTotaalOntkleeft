package TellerTotaalOntkleeft;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Objects;

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

    public static void setStyle(boolean bold, XSSFSheet totalSheet, int rowNumber, XSSFWorkbook XLSXWorkbookTotals){
        for(int i = 0; i < 2; i++){
            XSSFCellStyle styleOld = totalSheet.getRow(rowNumber).getCell(i).getCellStyle();
            XSSFCellStyle newStyle = XLSXWorkbookTotals.createCellStyle();
            newStyle.cloneStyleFrom(styleOld);
            newStyle.setBorderTop(BorderStyle.THIN);
            newStyle.setBorderBottom(BorderStyle.THIN);
            newStyle.setBorderLeft(BorderStyle.THIN);
            newStyle.setBorderRight(BorderStyle.THIN);
            //Font oldFont = totalSheet.getRow(rowNumber).getCell(i).getCellStyle().getFont();

            XSSFFont newFont = XLSXWorkbookTotals.createFont();
            newFont.setBold(bold);
            newFont.setFontName(newStyle.getFont().getFontName());
            newFont.setFamily(newStyle.getFont().getFamily());
            newStyle.setFont(newFont);

            totalSheet.getRow(rowNumber).getCell(i).setCellStyle(newStyle);
        }
    }
}
