package com.company;

import com.sun.org.apache.xerces.internal.parsers.IntegratedParserConfiguration;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Iterator;

public class ReadExcel {

    private static boolean checkIfRowIsEmpty(Row row) {
        if (row == null || row.getLastCellNum() <= 0) {
            return true;
        }
        Cell cell = row.getCell((int) row.getFirstCellNum());
        try {
            if (cell == null || "".equals(cell.getRichStringCellValue().getString())) {
                return true;
            }
        }catch (IllegalStateException ignored) {
            if (cell == null || "".equals(Double.toString(cell.getNumericCellValue()))) {
                return true;
            }
        }


        return false;

    }

    public static int emptyRowIndex(XSSFSheet sheet){
        Iterator<Row> iterator = sheet.iterator();
        int r = 0;
        while(iterator.hasNext()){

            Row row = iterator.next();
            if (!checkIfRowIsEmpty(row)) {
                r++;
            }

        }
       // System.out.println("ROWROWROWROOOW: "+r);
        return r;
    }
}
