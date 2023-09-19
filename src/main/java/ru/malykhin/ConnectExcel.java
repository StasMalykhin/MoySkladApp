package ru.malykhin;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

public class ConnectExcel {

    //Индикатор созданных заголовков
    private static boolean createdHeaders = false;

    public static String readFromExcel(File inputFile, int currentRow) throws IOException {

        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(inputFile));
        XSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);
        XSSFRow row = myExcelSheet.getRow(currentRow);

        //Читаем пока строка не пустая
        if (row != null) {
            XSSFCell firstCell = row.getCell(0);

            if (firstCell != null && firstCell.getCellType() == CellType.STRING) {
                return firstCell.getStringCellValue();
            }
        }
        myExcelBook.close();

        return null;
    }


    public static void writeIntoExcel(File inputFile, int nextRow, String nameProduct, Map<String, String> stocks,
                                      Map<String, String> quantityInStocks) throws IOException {

        //Создаем заголовки, если они еще не созданы
        if (!createdHeaders) {
            createHeaders(inputFile, stocks);
            createdHeaders = true;
        }

        Workbook book = new XSSFWorkbook(new FileInputStream(inputFile));
        Sheet sheet = book.getSheet("Наличие на складах");

        Row row = sheet.createRow(nextRow);
        Cell product = row.createCell(0);
        product.setCellValue(nameProduct);

        int cellNumber = 1;
        for (String stockWithRemnant : quantityInStocks.keySet()) {

            for (String nameStock : stocks.keySet()) {

                if (stockWithRemnant == nameStock) {
                    Cell remnantOnStock = row.createCell(cellNumber);
                    remnantOnStock.setCellValue(quantityInStocks.get(stockWithRemnant));
                }
                cellNumber++;
            }
            cellNumber = 1;
        }

        // Меняем размер столбца
        sheet.autoSizeColumn(1);

        // Записываем всё в файл
        book.write(new FileOutputStream(inputFile));
        book.close();
    }


    public static void createHeaders(File inputFile, Map<String, String> stocks) throws IOException {
        Workbook book = new XSSFWorkbook(new FileInputStream(inputFile));
        Sheet sheet = book.createSheet("Наличие на складах");

//Заполним заголовки таблицы

        Row row = sheet.createRow(0);
        Cell name = row.createCell(0);
        name.setCellValue("Наименование изделия");


        int cellNumberForStore = 1;

        for (String key : stocks.keySet()) {
            Cell store = row.createCell(cellNumberForStore);
            store.setCellValue(key);
            cellNumberForStore++;
        }


        // Меняем размер столбца
        for (int i = 0; i < stocks.size(); i++)
            sheet.autoSizeColumn(i);

        book.write(new FileOutputStream(inputFile));
        book.close();
    }
}
