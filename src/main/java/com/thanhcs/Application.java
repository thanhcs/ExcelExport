package com.thanhcs;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Application {
    public static void main(String[] args) {
        Application app = new Application();
        OPCPackage pkg = null;
        InputStream inputStream = null;
        try {
            System.out.println(app.getClass());
            inputStream = app.getClass().getClassLoader().getResourceAsStream("template.xlsx");
            pkg = OPCPackage.open(inputStream);
            XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            XSSFSheet firstSheet = workbook.getSheetAt(0);
            System.out.println("Number of row currently: " + firstSheet.getPhysicalNumberOfRows());
            System.out.println("Number of column currently: " + firstSheet.getRow(0).getPhysicalNumberOfCells());

            Object[][] datas = {
                    {"A", "Nguyen", new SimpleDateFormat("MM/dd/yyyy").parse("12/5/1993"), 22, "Computer Science", null},
                    {"B", "McCord", new SimpleDateFormat("MM/dd/yyyy").parse("5/5/2010"), 16, "NA", null},
                    {"Alice", "Tran", new SimpleDateFormat("MM/dd/yyyy").parse("1/7/1983"), 30, "Biology", null},
                    {"Peter", "Pan", new SimpleDateFormat("MM/dd/yyyy").parse("2/12/1989"), 27, "Biology", null},
            };
            int sampleRowIndex = firstSheet.getPhysicalNumberOfRows() - 1;
            Row sampleRow = firstSheet.getRow(sampleRowIndex);

            int rowNum = firstSheet.getPhysicalNumberOfRows();
            System.out.println("Creating excel");

            for (Object[] data : datas) {
                Row row = firstSheet.createRow(rowNum);
                int colNum = 0;
                for (Object field : data) {
                    Cell cell = row.createCell(colNum);
                    cell.setCellStyle(sampleRow.getCell(colNum).getCellStyle());
                    if (field == null) {
                        if (sampleRow.getCell(colNum).getCellTypeEnum() == CellType.FORMULA) {
                            String formula = sampleRow.getCell(colNum).getCellFormula();
                            String relativeFormula = formula.replaceAll("([A-Z]+)(" + (sampleRowIndex + 1) + ")", "$1" + (rowNum + 1)); //+1 due to the excel UI index
                            cell.setCellFormula(relativeFormula);
                        }
                        colNum++;
                    }
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    } else if (field instanceof Date) {
                        cell.setCellValue((Date) field);
                    }
                    colNum++;
                }
                rowNum++;
            }

            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

            // Removing sample row
            removeRow(firstSheet, sampleRowIndex);

            try {
                FileOutputStream outputStream = new FileOutputStream("out.xlsx");
                workbook.write(outputStream);
                workbook.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }

        } catch (InvalidFormatException e) {
            System.out.println("InvalidFormatException " + e.getMessage());
        } catch (IOException e) {
            System.out.println("IOException " + e.getMessage());
        } catch (ParseException e) {
            System.out.println("ParseException " + e.getMessage());
        } finally {
            try {
                if (pkg != null) {
                    pkg.close();
                }
                if (inputStream != null) {
                    inputStream.close();
                }
            } catch (IOException e) {
                System.out.println("IOException when trying to close " + e.getMessage());
            }
        }
        System.out.println("End.");
    }

    // Source: https://stackoverflow.com/questions/21946958/how-to-remove-a-row-using-apache-poi/21947170
    private static void removeRow(XSSFSheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        }
        if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }
}
