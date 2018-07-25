package de.haw.hamburg.haeckel.workinghoursedit;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.*;

import java.util.HashMap;
import java.util.Map;

public class WorkingHoursEntry {
    private static final Map<String, Integer> CellIndexOf = createMap();
    private static Map<String, Integer> createMap() {
        Map<String, Integer> result = new HashMap<String, Integer>();
        result.put("date", 0);
        result.put("start", 1);
        result.put("end", 2);
        result.put("duration", 3);
        result.put("task", 4);
        return result;
    }


    public static final double DEFAULT_HOURS_PER_WEEK = 39.0;
    public static final double DEFAULT_NUM_WORKING_DAYS = 5.0;
    public static final String [] header = {"Datum", "Start", "Ende", "Dauer", "Soll", "Überstunden", "Aufgabe"};

    private static final double hoursPerDay = DEFAULT_HOURS_PER_WEEK/DEFAULT_NUM_WORKING_DAYS;
    private final XSSFRow row;
    private int rowIdx;
    private boolean isHeader;

    //Styles
    private final XSSFCellStyle dateStyle;
    private final XSSFCellStyle timeStyle;
    private final XSSFCellStyle durationStyle;


    public WorkingHoursEntry(XSSFRow row, int rowIdx, boolean isHeader, XSSFWorkbook wb) {
        this.row = row;
        this.rowIdx = rowIdx;
        this.isHeader = isHeader;

        //date style
        dateStyle = wb.createCellStyle();
        CreationHelper createHelper = wb.getCreationHelper();
        dateStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("dd/mm/yy"));

        //time style
        timeStyle = wb.createCellStyle();
        timeStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("hh:mm"));

        //duration style
        durationStyle = wb.createCellStyle();
        durationStyle.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.00"));
    }

    public XSSFRow getRow() {
        return row;
    }

    public void fillRow(XSSFRow out) {
        if(isHeader){
            fillHeader(out);
        } else {
            fillData(out);
        }
    }

    private void fillData(XSSFRow out) {
        for (int i = 0; i < header.length; i++) {
            XSSFCell cell = out.createCell(i);

            switch (i){
                case 0:
                    fillCellDate(cell);
                    break;
                case 1:
                    fillCellStart(cell);
                    break;
                case 2:
                    fillCellEnd(cell);
                    break;
                case 3:
                    fillCellDuration(cell);
                    break;
                case 4:
                    fillCellHPD(cell);
                    break;
                case 5:
                    fillCellOvertime(cell);
                    break;
                case 6:
                    fillCellTask(cell);
                    break;
            }
        }
    }

    private void fillCellTask(XSSFCell cell) {
        cell.setCellType(XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(row.getCell(CellIndexOf.get("task")).getStringCellValue());
    }

    private void fillCellOvertime(XSSFCell cell) {
        String formula = "D" + (rowIdx+1) + "-E" + (rowIdx+1);
        cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
        cell.setCellFormula(formula);
        cell.setCellStyle(durationStyle);
    }

    private void fillCellHPD(XSSFCell cell) {
        cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellValue(hoursPerDay);
        cell.setCellStyle(durationStyle);
    }

    private void fillCellDuration(XSSFCell cell) {
        cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
        String str = row.getCell(CellIndexOf.get("duration")).getStringCellValue();
        str = str.replace("Std.", ":");
        str = str.replace("Min." , "");
        str = str.replace(" ", "");
        String [] split = str.split(":");
        double hours=0, minutes=0;
        if(split.length >= 1){
            hours = Double.parseDouble(split[0]);
        }
        if(split.length >= 2){
            minutes = Double.parseDouble(split[1])/60;
        }
        double value = hours + minutes;
        cell.setCellValue(value);
        cell.setCellStyle(durationStyle);
    }

    private void fillCellEnd(XSSFCell cell) {
        cell.setCellType(XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(row.getCell(CellIndexOf.get("end")).getStringCellValue());
        cell.setCellStyle(timeStyle);
    }

    private void fillCellStart(XSSFCell cell) {
        cell.setCellType(XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(row.getCell(CellIndexOf.get("start")).getStringCellValue());
        cell.setCellStyle(timeStyle);
    }

    private void fillCellDate(XSSFCell cell) {
        cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellValue(row.getCell(CellIndexOf.get("date")).getDateCellValue());
        cell.setCellStyle(dateStyle);
    }

    private void fillHeader(XSSFRow out) {
        for (int i = 0; i < header.length; i++) {
            XSSFCell cell = out.createCell(i);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(header[i]);
        }

    }

    @Override
    public String toString() {
        return "WorkingHoursEntry{" +
                "row=" + row +
                '}';
    }
}
