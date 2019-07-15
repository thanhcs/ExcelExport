package com.thanhcs;

import org.apache.commons.beanutils.PropertyUtils;
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
import java.lang.reflect.InvocationTargetException;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
            int numCols = firstSheet.getRow(0).getPhysicalNumberOfCells();
            System.out.println("Number of column currently: " + numCols);

            UserService userService = new UserService();
            List<ExportModel> datas = userService.getUsers();

            int sampleRowIndex = firstSheet.getLastRowNum() - 1;
            Row sampleRow = firstSheet.getRow(sampleRowIndex);

            int propertyNamesRowIndex = firstSheet.getLastRowNum();
            Row propertyNamesRow = firstSheet.getRow(propertyNamesRowIndex);
            Map<Integer, String> ColumnToPropertyMap = buildColumnToPropertyMapper(propertyNamesRow);

            int rowNum = firstSheet.getPhysicalNumberOfRows();
            System.out.println("Creating excel");

            for (ExportModel model : datas) {
                Row row = firstSheet.createRow(rowNum);
                int colNum = 0;
                for (int i = 0; i < numCols; ++i) {
                    Cell cell = row.createCell(colNum);
                    cell.setCellStyle(sampleRow.getCell(colNum).getCellStyle());
                    if (!ColumnToPropertyMap.containsKey(i)) {
                        if (sampleRow.getCell(colNum).getCellTypeEnum() == CellType.FORMULA) {
                            String formula = sampleRow.getCell(colNum).getCellFormula();
                            String relativeFormula = formula.replaceAll("([A-Z]+)(" + (sampleRowIndex + 1) + ")", "$1" + (rowNum + 1)); //+1 due to the excel UI index
                            cell.setCellFormula(relativeFormula);
                        }
                        colNum++;
                        continue;
                    }
                    Object field = PropertyUtils.getProperty(model, ColumnToPropertyMap.get(i));
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

            // Removing sample row
            removeRow(firstSheet, propertyNamesRowIndex);
            removeRow(firstSheet, sampleRowIndex);

            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

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
//        } catch (ParseException e) {
//            System.out.println("ParseException " + e.getMessage());
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
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

    private static Map<Integer, String> buildColumnToPropertyMapper(Row propertyNamesRow) {
        Map<Integer, String> mapper = new HashMap<>();
        for (int i = 0; i < propertyNamesRow.getLastCellNum(); ++i) {
            String cellValue = propertyNamesRow.getCell(i).getStringCellValue();
            if (cellValue != null) {
                mapper.put(i, cellValue);
            }
        }
        return mapper;
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

    //http://apache-poi.1045710.n5.nabble.com/Difference-of-getLastRowNum-and-getPhysicalNumberOfRows-td5723176.html
//    private static void addNewSection(XSSFSheet sheet, String[][] headers) {
//        int lastColumnIndex = getSampleColumn(sheet.getRow(0))
//
//    }

//    private static void addingColumn(String searchKey, XSSFSheet sheet, List<String[]> headers, boolean[] mergeHeaders) {
//        int sampleColumnIndex = getSampleColumn(sheet.getRow(0), searchKey);
//        if (sampleColumnIndex == -1) {
//            System.out.println("No Sample column. Stop.");
//        }
//        List<Cell> sampleCell = new ArrayList<Cell>();
//        IntStream.range(0, headers.size()).forEach(rowIndex -> sampleCell.add(sheet.getRow(rowIndex).getCell(sampleColumnIndex)));
//
//        int startPosition = sheet.getRow(0).getLastCellNum() - 1;
//        int numColumn = headers.get(0).length;
//        for (int i = 0; i < headers.size(); ++i) {
//            if (mergeHeaders[i]) {
//                int mergedCellIndex = sheet.addMergedRegion(new CellRangeAddress(i, i, startPosition, startPosition + numColumn - 1));
//                Cell mergeCell = sheet.getRow(0).createCell(mergedCellIndex);
//                mergeCell.setCellStyle(sampleCell.get(0).getCellStyle());
//                mergeCell.setCellStyle(sampleCell.get(0).getCellStyle());
//                mergeCell.setCellValue(headers.get(i)[0]);
//            } else {
//                String[] headerTitles = headers.get(i);
//                int tempStartPosition = startPosition;
//                for (String title : headerTitles) {
//                    CellStyle cellStyle = sampleCell.get(i).getCellStyle();
//                    Cell headerCell = sheet.getRow(i).createCell(tempStartPosition);
//                    headerCell.setCellStyle(cellStyle);
//                    headerCell.setCellValue(title);
//                    ++tempStartPosition;
//                }
//            }
//        }
//        System.out.println("Finish adding new session with searchKey: " + searchKey);
//    }

    private static int getSampleColumn(Row row, String searchKey) {
        for (int i = 0; i < row.getLastCellNum(); ++i) {
            if (row.getCell(i).getStringCellValue().equals(searchKey)) {
                return i;
            }
        }
        return -1;
    }
}
