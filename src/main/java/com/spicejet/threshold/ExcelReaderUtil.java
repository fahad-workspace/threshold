package com.spicejet.threshold;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelReaderUtil {

    private String getCellValue(Cell cell) {
        if (cell != null) {
            cell.setCellType(CellType.STRING);
            return cell.getStringCellValue();
        }
        return "";
    }

    public List<List<Object>> readBooksFromExcelFile(String excelFilePath, int count) throws IOException {
        ArrayList<List<Object>> listBooks = new ArrayList<>();
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        for (Row nextRow : firstSheet) {
            List<Object> aBook = new ArrayList<>();
            for (int i = 0; i < count - 1; i++) {
                Cell nextCell = nextRow.getCell(i);
                aBook.add(getCellValue(nextCell));
            }
            if (nextRow.getRowNum() == 0) {
                aBook.add("Type");
                aBook.add("Base");
                aBook.add("THR Base Dim");
                aBook.add("THR Base Amount");
                aBook.add("Dim"); // dim
                aBook.add("Amount"); // dim amount
                listBooks.add(aBook);
                continue;
            }
            String second;
            // dim 1
            String first = getCellValue(nextRow.getCell(count + 2));
            if (first != null && first.length() > 0) {
                List<Object> bBook = new ArrayList<>();
                bBook.addAll(aBook);
                second = getCellValue(nextRow.getCell(count + 3));
                if (first.equalsIgnoreCase("H")) {
                    bBook.add("B"); // type
                    bBook.add(""); // base
                    bBook.add(""); // THR base dim
                    bBook.add(""); // THR base amount
                    bBook.add(first); // dim
                    bBook.add(second); // dim amount
                }
                if (first.equalsIgnoreCase("AH")) {
                    bBook.add("B"); // type
                    bBook.add(""); // base
                    bBook.add(""); // THR base dim
                    bBook.add(""); // THR base amount
                    bBook.add(first); // dim
                    bBook.add(second); // dim amount
                }
                listBooks.add(bBook);
            }
            // dim 2
            first = getCellValue(nextRow.getCell(count + 4));
            if (first != null && first.length() > 0) {
                List<Object> bBook = new ArrayList<>();
                bBook.addAll(aBook);
                second = getCellValue(nextRow.getCell(count + 5));
                if (first.equalsIgnoreCase("C")) {
                    bBook.add("B"); // type
                    bBook.add(""); // base
                    bBook.add(""); // THR base dim
                    bBook.add(""); // THR base amount
                    bBook.add(first); // dim
                    bBook.add(second); // dim amount
                }
                if (first.equalsIgnoreCase("AH")) {
                    bBook.add("B"); // type
                    bBook.add(""); // base
                    bBook.add(""); // THR base dim
                    bBook.add(""); // THR base amount
                    bBook.add(first); // dim
                    bBook.add(second); // dim amount
                }
                listBooks.add(bBook);
            }
            // dim 3
            first = getCellValue(nextRow.getCell(count + 6));
            if (first != null && first.length() > 0) {
                List<Object> bBook = new ArrayList<>();
                bBook.addAll(aBook);
                second = getCellValue(nextRow.getCell(count + 7));
                if (first.equalsIgnoreCase("D") || first.equalsIgnoreCase("MT") ||
                        first.equalsIgnoreCase("YR")) {
                    String base = getCellValue(nextRow.getCell(count - 1));
                    if (base != null && base.length() > 0) {
                        if (base.equalsIgnoreCase("E") || base.equalsIgnoreCase("M") ||
                                base.equalsIgnoreCase("D") || base.equalsIgnoreCase("I") ||
                                base.equalsIgnoreCase("S")) {
                            bBook.add("W"); // type
                            bBook.add(base); // base
                            bBook.add(""); // THR base dim
                            bBook.add(""); // THR base amount
                            bBook.add(first); // dim
                            bBook.add(second); // dim amount
                        }
                        if (base.equalsIgnoreCase("V")) {
                            bBook.add("W"); // type
                            bBook.add(base); // base
                            bBook.add(getCellValue(nextRow.getCell(count))); // THR base dim
                            bBook.add(getCellValue(nextRow.getCell(count + 1))); // THR base amount
                            bBook.add(first); // dim
                            bBook.add(second); // dim amount
                        }
                    }
                }
                if (first.equalsIgnoreCase("DT")) {
                    bBook.add("B"); // type
                    bBook.add(""); // base
                    bBook.add(""); // THR base dim
                    bBook.add(""); // THR base amount
                    bBook.add("D"); // dim
                    bBook.add(second); // dim amount
                }
                listBooks.add(bBook);
            }
            // idim 1
            first = getCellValue(nextRow.getCell(count + 8));
            if (first != null && first.length() > 0) {
                List<Object> bBook = new ArrayList<>();
                bBook.addAll(aBook);
                second = getCellValue(nextRow.getCell(count + 9));
                if (first.equalsIgnoreCase("H")) {
                    bBook.add("I"); // type
                    bBook.add(""); // base
                    bBook.add(""); // THR base dim
                    bBook.add(""); // THR base amount
                    bBook.add(first); // dim
                    bBook.add(second); // dim amount
                }
                if (first.equalsIgnoreCase("AH")) {
                    bBook.add("I"); // type
                    bBook.add(""); // base
                    bBook.add(""); // THR base dim
                    bBook.add(""); // THR base amount
                    bBook.add(first); // dim
                    bBook.add(second); // dim amount
                }
                listBooks.add(bBook);
            }
            // idim 2
            first = getCellValue(nextRow.getCell(count + 10));
            if (first != null && first.length() > 0) {
                List<Object> bBook = new ArrayList<>();
                bBook.addAll(aBook);
                second = getCellValue(nextRow.getCell(count + 11));
                if (first.equalsIgnoreCase("C")) {
                    bBook.add("I"); // type
                    bBook.add(""); // base
                    bBook.add(""); // THR base dim
                    bBook.add(""); // THR base amount
                    bBook.add(first); // dim
                    bBook.add(second); // dim amount
                }
                if (first.equalsIgnoreCase("AC")) {
                    bBook.add("I"); // type
                    bBook.add(""); // base
                    bBook.add(""); // THR base dim
                    bBook.add(""); // THR base amount
                    bBook.add(first); // dim
                    bBook.add(second); // dim amount
                }
                listBooks.add(bBook);
            }
            // idim 3
            first = getCellValue(nextRow.getCell(count + 12));
            if (first != null && first.length() > 0) {
                List<Object> bBook = new ArrayList<>();
                bBook.addAll(aBook);
                second = getCellValue(nextRow.getCell(count + 13));
                if (first.equalsIgnoreCase("D")) {
                    bBook.add("I"); // type
                    bBook.add(""); // base
                    bBook.add(""); // THR base dim
                    bBook.add(""); // THR base amount
                    bBook.add(first); // dim
                    bBook.add(second); // dim amount
                }
                listBooks.add(bBook);
            }
        }
        workbook.close();
        inputStream.close();
        return listBooks;
    }

    public void writeBooksFromExcelFile(String inputExcelFilePath, String outputExcelFilePath, List<List<Object>> listBooks)
            throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet(getSheetName(inputExcelFilePath));
            int rowNum = 0;
            System.out.println("Processing : " + inputExcelFilePath
                    .split(Constants.SEPERATOR)[inputExcelFilePath.split(Constants.SEPERATOR).length - 1]);
            for (List<Object> book : listBooks) {
                if (rowNum == 0) {
                    rowNum++;
                    createMainHeader(sheet, workbook, book);
                    continue;
                }
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (Object aBook : book) {
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(aBook.toString());
                }
            }
            FileOutputStream outputStream = new FileOutputStream(outputExcelFilePath);
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Processed : " + inputExcelFilePath
                .split(Constants.SEPERATOR)[inputExcelFilePath.split(Constants.SEPERATOR).length - 1]);
    }

    private String getSheetName(String excelFilePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        workbook.close();
        return firstSheet.getSheetName();
    }

    private void createMainHeader(XSSFSheet sheet, XSSFWorkbook workbook, List<Object> book) {
        Row header = sheet.createRow(0);
        for (int i = 0; i < book.size(); i++) {
            header.createCell(i).setCellValue(book.get(i).toString());
        }
        setHeaderStyle(workbook, header, sheet, IndexedColors.GREY_40_PERCENT.getIndex());
    }

    private void setHeaderStyle(XSSFWorkbook workbook, Row header, XSSFSheet sheet, short backgroundColor) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(backgroundColor);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        style.setTopBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        style.setRightBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        style.setLeftBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        for (int i = 0; i < header.getLastCellNum(); i++) {
            if (header.getCell(i) != null) {
                header.getCell(i).setCellStyle(style);
                sheet.autoSizeColumn(i);
            }
        }
    }
}
