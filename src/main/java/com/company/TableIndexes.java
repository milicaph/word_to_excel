package com.company;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class TableIndexes {

    private final String boardSeatsReference = "Board Seats";
    private final String[] currentHeader = {"Company", "Industry", "Ownership Status", "Financing Status",
            "Location", "Since", "Representing", "Firm Type"};
    private final String[] formerHeader = {"Company", "Industry", "Ownership Status", "Financing Status",
            "Location", "From", "Until"};


    private ArrayList<Integer> currentIndexes = new ArrayList<>();
    private ArrayList<Integer> formerIndexes = new ArrayList<>();

    private XWPFDocument getDocument(String path) throws IOException {
        System.out.println(path);
        File file = new File(path);
        FileInputStream fis = new FileInputStream(file);

        return new XWPFDocument(fis);

    }

    private List<XWPFTable> tables(XWPFDocument doc) {
        return doc.getTables();

    }

    private boolean getTableIndex(String cellText) {
        return cellText.contains(boardSeatsReference);

    }

    private boolean cellValueContains(String cellText, String tableRef) {
        return cellText.contains(tableRef);

    }

    private boolean docContains(XWPFDocument xdoc, String refString) throws IOException {
        return readWholeDocx(xdoc).contains(refString);

    }

    public static String readWholeDocx(XWPFDocument xdoc) throws IOException {
        String wholeDocx = "";

        try {
            XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
            wholeDocx = extractor.getText();

            return wholeDocx;

        } catch (Exception ex) {
            ex.printStackTrace();
        }

        return wholeDocx;
    }

    public void getTableIndexes(String path) throws IOException {
        XWPFDocument doc = getDocument(path);
        List<XWPFTable> tables = tables(doc);

        //File testFile = new File(path);
        FileInputStream fis = new FileInputStream("src/main/resources/Table indexes.xlsx");

        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        if (!docContains(doc, boardSeatsReference)) {
            //currentIndexes.add(0);
            //formerIndexes.add(0);
            System.out.println(path + "---------------------- " + 0);
        } //else if(!docContains(doc, boardSeatsReference))

        int t = 0;

        for (XWPFTable table : tables) {
            //System.out.println("Table: " + t);

            int er = ReadExcel.emptyRowIndex(sheet);
            int r = 0;
            XSSFRow rowX = sheet.createRow(er);
            for (XWPFTableRow row : table.getRows()) {
                //System.out.println("Row: " + r + " " +row.getTableCells().);

                int c = 0;
                for (XWPFTableCell cell : row.getTableCells()) {

                    String sFieldValue = cell.getText();

                    if (r == 0 & c == 6) {
                        if (cellValueContains(sFieldValue, currentHeader[6])) {
                            XSSFCell cell0 = rowX.createCell(0);
                            cell0.setCellValue(t);
                            XSSFCell cell2 = rowX.createCell(2);
                            cell2.setCellValue(path);
                            //currentIndexes.add(t);
                            System.out.println(path + "---------------------- " + t);
                        } else if (cellValueContains(sFieldValue, formerHeader[6])) {
                            XSSFCell cell1 = rowX.createCell(1);
                            cell1.setCellValue(t);
                            XSSFCell cell2 = rowX.createCell(2);
                            cell2.setCellValue(path);
                            //formerIndexes.add(t);
                            System.out.println(path + "---------------------- " + t);
                        }
                    }

                    c++;

                }

                r++;
                System.out.println(" ");

            }
            t++;

        }

        try {
            FileOutputStream outputStream = new FileOutputStream("src/main/resources/Table indexes.xlsx");
            workbook.write(outputStream);
        } catch (IndexOutOfBoundsException | IOException e) {
            e.printStackTrace();
        }


    }

    public void writeTableData() throws IOException {
        String path;

        FileInputStream fisRead = new FileInputStream("src/main/resources/Table indexes.xlsx");
        XSSFWorkbook workbookRead = new XSSFWorkbook(fisRead);
        XSSFSheet sheet = workbookRead.getSheet("Sheet1");
        Iterator<Row> iterator = sheet.rowIterator();
        int i = 0;
        while (iterator.hasNext()) {
            Row row = iterator.next();
            Cell cell0 = row.getCell(0);
            Cell cell1 = row.getCell(1);
            Cell cell2 = row.getCell(2);

            int currInt = (int)cell0.getNumericCellValue();
            int forInt = (int)cell1.getNumericCellValue();
            path = cell2.getStringCellValue();

            XWPFDocument doc = getDocument(path);
            List<XWPFTable> tables = tables(doc);
            XWPFTable tableCurrent = tables.get(currInt);
            XWPFTable tableFormer = tables.get(forInt);

            FileInputStream fisWrite = new FileInputStream("src/main/resources/Board.xlsx");
            XSSFWorkbook workbookWrite = new XSSFWorkbook(fisWrite);
            XSSFSheet sheetCurrent = null;
            XSSFSheet sheetFormer = null;
            
           if(i == 0) {
               sheetCurrent = workbookWrite.createSheet("Current Board");
               sheetFormer = workbookWrite.createSheet("Former Board");
           } else if(i > 0){
               sheetCurrent = workbookWrite.getSheet("Current Board");
               sheetFormer = workbookWrite.getSheet("Former Board");
            }
           
            String pbid = getPBID(doc);

            writeOutput(tableCurrent, sheetCurrent, pbid);
            writeOutput(tableFormer, sheetFormer, pbid);

            i++;

        try {
            FileOutputStream outputStream = new FileOutputStream("src/main/resources/Board.xlsx");
            workbookWrite.write(outputStream);
        } catch (IndexOutOfBoundsException | IOException e) {
            e.printStackTrace();
        }

        }
    }

        private static String getPBID (XWPFDocument doc) throws IOException {
            String text = readWholeDocx(doc).toLowerCase();
            String endValue = "", indexBeg = "pbid",
                    indexEnd = "general information",
                    //indexBegAlt = indexEnd,
                    indexEndAlt = "biography";

            int i = text.indexOf(indexBeg) + indexBeg.length() + 2;
            int ib = text.indexOf(indexEnd);
            int ic = text.indexOf(indexEndAlt);
            //int id = ic;

            if (!text.contains(indexBeg)) {

                return endValue;
            }


            if (i < ib) {

                try {
                    endValue = text.substring(i, ib)
                            .trim();
                } catch (Exception e) {
                    e.fillInStackTrace();
                }

            } else if (i > ib) {
                try {
                    endValue = text.substring(i, ic)
                            .trim();
                } catch (Exception e) {
                    e.fillInStackTrace();
                }

                int ie = endValue.indexOf("-") + 4;
                endValue = endValue.substring(0, ie);

                return endValue;
            }

            return endValue.substring(0, endValue.indexOf("p") + 1);

        }

        private static void writeOutput (XWPFTable table, XSSFSheet sheet, String pbid){
            int sr = 0;

            for (XWPFTableRow row : table.getRows()) {

                System.out.println(sr);
                int r = ReadExcel.emptyRowIndex(sheet);
                XSSFRow rowX = sheet.createRow(r);
                StringBuilder str = new StringBuilder();

                int sc = 0;
                for (XWPFTableCell cell : row.getTableCells()) {


                    if (sc == 0) {
                        XSSFCell cell0 = rowX.createCell(sc);
                        cell0.setCellValue(pbid);
                        XSSFCell cell1 = rowX.createCell(sc + 1);
                        String sFieldValue = cell.getText();
                        cell1.setCellValue(sFieldValue);
                    } else {
                        XSSFCell cell0 = rowX.createCell(sc + 1);
                        String sFieldValue = cell.getText();
                        cell0.setCellValue(sFieldValue);
                    }
                    sc++;
                }
                sr++;

            }


        }

    }









