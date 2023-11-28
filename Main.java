import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
        readExcelFile("a.xlsx", 0, 2);  // Пример: выводим только первые три столбца (индексы 0, 1, 2)
    }

    private static void readExcelFile(String filePath, int startColumn, int endColumn) {
        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
            processRows(sheet, formulaEvaluator, startColumn, endColumn);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void processRows(XSSFSheet sheet, FormulaEvaluator formulaEvaluator, int startColumn, int endColumn) {
        for (Row row : sheet) {
            processCells(row, formulaEvaluator, startColumn, endColumn);
            System.out.println();  // Новая строка после обработки каждой строки
        }
    }

    private static void processCells(Row row, FormulaEvaluator formulaEvaluator, int startColumn, int endColumn) {
        int[] columnWidths = {50, 20, 30};  // Пример: каждый столбец имеет свою ширину
    
        for (int colIndex = startColumn; colIndex <= endColumn; colIndex++) {
            Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            String cellValue = getCellValue(cell, formulaEvaluator);
    
            // Добавление табуляции перед каждой ячейкой, кроме первой
            if (colIndex > startColumn) {
                System.out.print("\t");
            }
    
            // Форматирование строки с заданной шириной столбца
            System.out.print(String.format("%-" + columnWidths[colIndex - startColumn] + "s", cellValue));
        }
        System.out.println();  // Новая строка после обработки каждой строки
    }
    
    
private static String getCellValue(Cell cell, FormulaEvaluator formulaEvaluator) {
    switch (cell.getCellTypeEnum()) {
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                // Если ячейка содержит дату, возвращаем ее в формате строки
                return new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
            } else {
                // Возвращаем числовое значение ячейки
                return String.valueOf(cell.getNumericCellValue());
            }
        case STRING:
            // Возвращаем строковое значение ячейки
            return cell.getStringCellValue();
        default:
            return cell.getCellTypeEnum() == CellType.FORMULA
                    ? evaluateFormulaCell(cell, formulaEvaluator)
                    : cell.getRichStringCellValue().getString();
    }
}

    private static String evaluateFormulaCell(Cell cell, FormulaEvaluator formulaEvaluator) {
        // Вычисление значения формулы в ячейке
        CellValue cellValue = formulaEvaluator.evaluate(cell);
        return String.valueOf(cellValue.getNumberValue());
    }
}
