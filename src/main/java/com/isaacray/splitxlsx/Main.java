package com.isaacray.splitxlsx;

import com.monitorjbl.xlsx.StreamingReader;
import com.monitorjbl.xlsx.exceptions.NotSupportedException;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.StopWatch;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Main {

    public static void main(String[] argsv) throws Exception{
        //This can go in src/main/resources
        String fileName = "sbp_locations_full_dedup.xlsx";
        int rowsPerFile = 10000;

        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        System.out.println("Removing old split files...");
        runCommand("rm -Rfv splitFile*");
        System.out.println("Removing old tar files...");
        runCommand("rm -Rfv *.tgz");

        Workbook excelWorkbook = createExcelWorkbook(new File(ClassLoader.getSystemResource(fileName).getFile()));
        Iterator<Row> rowIterator = excelWorkbook.getSheetAt(0).rowIterator();
        long totalRows = 0;
        System.out.println("Counting total rows...");
        stopWatch.split();
        System.out.println("stopWatch = " + stopWatch.toSplitString());
        while(rowIterator.hasNext()) {
            rowIterator.next();
            totalRows++;
        }
        System.out.println("totalRows = " + totalRows);
        stopWatch.split();
        System.out.println("stopWatch = " + stopWatch.toSplitString());
        excelWorkbook.close();

        excelWorkbook = createExcelWorkbook(new File(ClassLoader.getSystemResource(fileName).getFile()));
        rowIterator = excelWorkbook.getSheetAt(0).rowIterator();
        Row header = rowIterator.next();
        int fileCounter = 0;
        long loopCounter = 0;

        XSSFWorkbook workbook = new XSSFWorkbook();
        workbook.createSheet();
        addRow(workbook, header, header, 0);

        while(rowIterator.hasNext()){

            addRow(workbook, header, rowIterator.next(), 1);

            if(loopCounter == rowsPerFile){
                loopCounter = 0;
                workbook.write(new FileOutputStream("splitFile" + StringUtils.leftPad(String.valueOf(fileCounter), 3, "0") + ".xlsx"));
                workbook = new XSSFWorkbook();
                workbook.createSheet();
                addRow(workbook, header, header, 0);
                fileCounter++;
                System.out.print("fileCounter = " + fileCounter + " | ");
                Double percent = Double.valueOf(fileCounter * rowsPerFile)/totalRows * 100;
                System.out.print(Math.round(percent) + "% complete. | ");
                stopWatch.split();
                System.out.println("stopWatch = " + stopWatch.toSplitString());
            }
            ++loopCounter;
        }

        workbook.write(new FileOutputStream("splitFile" + StringUtils.leftPad(String.valueOf(fileCounter), 3, "0") + ".xlsx"));

        runCommand("tar -cvzf " + fileName.split("\\.")[0] + "_splitBy" + rowsPerFile + ".tgz splitFile*");
    }

    private static void runCommand(String command) throws Exception{
        String[] cmd = {"/bin/sh", "-c", command};
        Process process = Runtime.getRuntime().exec(cmd);
        process.waitFor();
        BufferedReader stdInput = new BufferedReader(new InputStreamReader(process.getInputStream()));
        String s;
        while ((s = stdInput.readLine()) != null) {
            System.out.println(s);
        }

        stdInput = new BufferedReader(new InputStreamReader(process.getErrorStream()));
        while ((s = stdInput.readLine()) != null) {
            System.out.println(s);
        }
    }

    private static void addRow(Workbook workbook, Row header, Row rowToAdd, int headerRow){
        Sheet sheetAt = workbook.getSheetAt(0);
        Row row = sheetAt.createRow(sheetAt.getLastRowNum() + headerRow);
        List<Cell> cells = new ArrayList();
        rowToAdd.cellIterator().forEachRemaining(cells::add);
        header.forEach(cell -> {
            Cell createdCell = row.createCell(cell.getColumnIndex());
            rowToAdd.cellIterator().hasNext();
            try{
                createdCell.setCellValue(cells.get(cell.getColumnIndex()).getStringCellValue());
            } catch (NotSupportedException notSupportedException){
                // empty cell
            }
        });
    }

    private static Workbook createExcelWorkbook(File file) throws IOException, InvalidFormatException {
        File temp = File.createTempFile("upload-", ".tmp");
        FileOutputStream outputStream = new FileOutputStream(temp);
        byte[] bytes = Files.readAllBytes(file.toPath());
        outputStream.write(bytes);
        outputStream.close();
        Workbook workbook;
        try {
            workbook = StreamingReader.builder()
                .rowCacheSize(100)
                .bufferSize(4096)
                .open(new FileInputStream(temp));
        } catch (OLE2NotOfficeXmlFileException e) {
            workbook = WorkbookFactory.create(temp);
        }
        return workbook;
    }

}
