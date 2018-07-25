package de.haw.hamburg.haeckel.workinghoursedit;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.*;

public class Processor {

    private final String inFile;

    private final String outFile;

    private XSSFWorkbook wbIn;

    private XSSFWorkbook wbOut;

    private Map<String, List<WorkingHoursEntry>> entrySheets = new HashMap<String, List<WorkingHoursEntry>>();


    public Processor(String inFile, String outFile) throws FileNotFoundException {
        this.inFile = inFile;
        this.outFile = outFile;

        // create a Blank Workbook
        wbOut = new XSSFWorkbook();

        // load input workbook
        openWbFromFile();

        // read entries
        readEntries();

        // write entries to outfile.
        writeEntries();

        // export
        writeWbToFile();
    }

    private void appendStatistics(XSSFSheet sheet) {
        int lastRow = sheet.getLastRowNum();

        //overtime
        XSSFRow otRow = sheet.createRow(lastRow+2);
        XSSFCell otTitleCell = otRow.createCell(4);
        otTitleCell.setCellType(XSSFCell.CELL_TYPE_STRING);
        otTitleCell.setCellValue("Überstunden Gesamt: ");
        XSSFCell otFormulaCell = otRow.createCell(5);
        otFormulaCell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
        otFormulaCell.setCellFormula("SUM(F2:F" + (lastRow+1) + ")");
        CreationHelper createHelper = wbOut.getCreationHelper();
        XSSFCellStyle durationStyle = wbOut.createCellStyle();
        durationStyle.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.00"));
        otFormulaCell.setCellStyle(durationStyle);

        //sick days
        int sickDays = getTaskCount(sheet,6,"krank");
        XSSFRow sRow = sheet.createRow(lastRow+3);
        XSSFCell sTitleCell = sRow.createCell(4);
        sTitleCell.setCellType(XSSFCell.CELL_TYPE_STRING);
        sTitleCell.setCellValue("Krankheitstage: ");
        XSSFCell sValueCell = sRow.createCell(5);
        sValueCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
        sValueCell.setCellValue(sickDays);

        //vacation days
        int vacDays = getTaskCount(sheet,6,"urlaub");
        XSSFRow vRow = sheet.createRow(lastRow+4);
        XSSFCell vTitleCell = vRow.createCell(4);
        vTitleCell.setCellType(XSSFCell.CELL_TYPE_STRING);
        vTitleCell.setCellValue("Urlaubstage: ");
        XSSFCell vValueCell = vRow.createCell(5);
        vValueCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
        vValueCell.setCellValue(vacDays);

    }

    private int getTaskCount(XSSFSheet sheet, int column, String task) {
        int count = 0;

        Iterator<Row> rowIt = sheet.iterator();

        while(rowIt.hasNext()){
            Row row = rowIt.next();
            if(row != null) {
                Cell cell = row.getCell(column);
                if (cell != null) {
                    String value = cell.getStringCellValue();

                    if (value.toUpperCase().contains(task.toUpperCase())) {
                    count++;
                    }
                }
            }
        }

        return count;
    }

    private void writeEntries() {
        System.out.println("Writing entries to new Workbook...");
        Set<Map.Entry<String, List<WorkingHoursEntry>>> sheetEntrySet = entrySheets.entrySet();

        for (Map.Entry<String, List<WorkingHoursEntry>> sheetEntry: sheetEntrySet){
            XSSFSheet sheet = wbOut.createSheet(sheetEntry.getKey());

            List<WorkingHoursEntry> entryList = sheetEntry.getValue();

            Iterator<WorkingHoursEntry> entries = entryList.iterator();
            for (int rowId = 0; entries.hasNext(); rowId++){
                WorkingHoursEntry entry = entries.next();
                XSSFRow row = sheet.createRow(rowId);

                entry.fillRow(row);
            }

            appendStatistics(sheet);

            autoFormatSheet(sheet);

        }

    }

    private void autoFormatSheet(XSSFSheet sheet) {

        // Define a Conditional Formatting rule, which triggers formatting
        // when cell's value is greater or equal than 100.0 and
        // applies patternFormatting defined below.
        XSSFSheetConditionalFormatting cf = sheet.getSheetConditionalFormatting();

        XSSFConditionalFormattingRule rule = cf.createConditionalFormattingRule(
                ComparisonOperator.LT,
                "0", // 1st formula
                null     // 2nd formula is not used for comparison operator

        );
        XSSFFontFormatting fontFmt = rule.createFontFormatting();
        fontFmt.setFontColorIndex(IndexedColors.DARK_RED.index);

        CellRangeAddress[] range = {
                new CellRangeAddress(1, sheet.getLastRowNum()+1, 5,5)
        };

        cf.addConditionalFormatting(range, rule);

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(4);
        sheet.autoSizeColumn(5);
        sheet.autoSizeColumn(6);
    }

    private void readEntries() {
        System.out.println("Reading entries from old Workbook...");
        Iterator<Sheet> sheets = wbIn.sheetIterator();
        //for all sheets in workbook
        for (int i = 0; i < wbIn.getNumberOfSheets(); i++) {

            XSSFSheet sheet = wbIn.getSheetAt(i);

            //create new entries list
            List<WorkingHoursEntry> entries = new ArrayList<WorkingHoursEntry>();

            for (int j = 0; j < sheet.getLastRowNum(); j++) {
                XSSFRow row = sheet.getRow(j);

                if(row != null){
                    WorkingHoursEntry entry;
                    if (j == 0){
                        entry = new WorkingHoursEntry(row, j, true, wbOut);
                    } else {
                        entry = new WorkingHoursEntry(row, j,false, wbOut);
                    }


                    entries.add(entry);
                }
            }


            entrySheets.put(sheet.getSheetName(), entries);
        }
    }


    public void writeWbToFile() throws FileNotFoundException {
        System.out.println("Exporting new Workbook...");
        //Create file system using specific name
        FileOutputStream os = new FileOutputStream(new File(outFile));

        //write operation workbook using file out object
        try {
            wbOut.write(os);
            os.close();
            System.out.println(outFile + " written successfully");
        } catch (IOException e) {
            System.err.println(outFile + " write error");
            e.printStackTrace();
        }
    }

    public void openWbFromFile() throws FileNotFoundException {
        System.out.println("Open old Workbook...");

        FileInputStream is = new FileInputStream(new File(inFile));

        //Get the workbook instance for XLSX file
        try {
            wbIn = new XSSFWorkbook(is);

            System.out.println(inFile + " import successful");

        } catch (IOException e) {
            System.err.println(inFile + " import failed");
            e.printStackTrace();
        }


    }
}
