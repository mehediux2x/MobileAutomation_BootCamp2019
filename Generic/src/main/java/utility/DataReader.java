package utility;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
public class DataReader  {

    public static XSSFWorkbook wb = null;
    public static XSSFSheet sheet = null;
    public static XSSFCell cell = null;
    public static FileOutputStream fio = null;
    public static int numberOfRows, numberOfCol, rowNum;

    public static String[][] fileReader1(String path,int sheetIndex) throws IOException {
        String[][] data = {};
        File file = new File(path);
        FileInputStream fis = new FileInputStream(file);
        wb = new XSSFWorkbook(fis);
        sheet = wb.getSheetAt(sheetIndex);
        numberOfRows = sheet.getLastRowNum();
        numberOfCol = sheet.getRow(0).getLastCellNum();
        data = new String[numberOfRows + 1][numberOfCol + 1];

        for (int i = 1; i < data.length; i++) {
            XSSFRow rows = sheet.getRow(i);
            for (int j = 0; j < numberOfCol; j++) {
                XSSFCell cell = rows.getCell(j);
                String cellData = getCellValue(cell);
                data[i][j] = cellData;
            }
        }
        return data;
    }

    public static String[] fileReader2(String path, int sheetIndex) throws IOException {
        String[] data = {};
        File file = new File(path);
        FileInputStream fis = new FileInputStream(file);
        wb = new XSSFWorkbook(fis);
        sheet = wb.getSheetAt(sheetIndex);
        numberOfRows = sheet.getLastRowNum();
        numberOfCol = sheet.getRow(0).getLastCellNum();
        data = new String[numberOfRows + 1];

        for (int i = 1; i < data.length; i++) {
            XSSFRow rows = sheet.getRow(i);
            for (int j = 0; j < numberOfCol; j++) {
                XSSFCell cell = rows.getCell(j);
                String cellData = getCellValue(cell);
                data[i] = cellData;
            }
        }
        return data;
    }

    public static String getCellValue(XSSFCell cell) {
        Object value = null;

        int dataType = cell.getCellType();
        switch (dataType) {
            case XSSFCell.CELL_TYPE_NUMERIC:
                value = cell.getNumericCellValue();
                break;
            case XSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
        }
        return value.toString();

    }

    public static void writeBack(String value) throws IOException {
        wb = new XSSFWorkbook();
        sheet = wb.createSheet();
        XSSFRow row = sheet.createRow(rowNum);
        row.setHeightInPoints(10);

        fio = new FileOutputStream(new File("ExcelFile.xls"));
        wb.write(fio);
        for (int i = 0; i < row.getLastCellNum(); i++) {
            row.createCell(i);
            cell.setCellValue(value);
        }
        fio.close();
        wb.close();
    }
}