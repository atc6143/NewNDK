package com.example.lib;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChartSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.omg.Messaging.SYNC_WITH_TRANSPORT;
import org.testng.IInvokedMethod;
import org.testng.IReporter;
import org.testng.ISuite;
import org.testng.ISuiteResult;
import org.testng.ITestContext;
import org.testng.ITestResult;
import org.testng.TestNG;
import org.testng.annotations.Test;
import org.testng.xml.XmlSuite;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

public class AutoTestReporter implements IReporter {


    public static String FILE_SEPARATOR = System.getProperty("file.separator");
    public static String FILE_PATH = "E:" + FILE_SEPARATOR + "Android_O_AUTO" + FILE_SEPARATOR + "MyAutoTest" + FILE_SEPARATOR +
            "lib" + FILE_SEPARATOR + "resource" + FILE_SEPARATOR + "sourceTestReport" + FILE_SEPARATOR + "33.xlsx";

    @Override
    public void generateReport(List<XmlSuite> xmlSuites, List<ISuite> suites, String outputDirectory) {

        Map<String, List<String>> testCaseMap = new TreeMap<>();
        Map<String, List<String>> startCaseMap = getExcelValue(FILE_PATH);
        for (ISuite s : suites) {
            for (ISuiteResult suiteResult : s.getResults().values()) {
                ITestContext context = suiteResult.getTestContext();
                List<IInvokedMethod> allInvokedMethods = context.getSuite().getAllInvokedMethods();
                for (IInvokedMethod method : allInvokedMethods) {
                    if (method.isTestMethod()) {
                        List<String> startCase = new ArrayList<>();
                        String methodName = method.getTestMethod().getMethodName();
                        startCase.add(methodName);
                        List<String> allValues = startCaseMap.get(methodName);
                        for (String s1 : allValues) {
                            System.out.println(s1);
                            startCase.add(s1);
                        }
                        String testResult = getStatusString(method.getTestResult().getStatus());
                        String resultComment = "";
                        startCase.add(testResult);
                        if (testResult.equals("PASS")) {
                            resultComment = "";
                        } else if (testResult.equals("FAIL")) {
                            resultComment = "FAIL";
                        } else if (testResult.equals("SKIP")) {
                            resultComment = "SKIP";
                        }
                        startCase.add(resultComment);
                        testCaseMap.put(methodName, startCase);
                    }
                }
            }
        }
        writeExcel("Sanity_test", FILE_PATH, testCaseMap);

    }

    public String getStatusString(int num) {
        switch (num) {
            case ITestResult.CREATED:
                return "SKIP";
            case ITestResult.SUCCESS:
                return "PASS";
            case ITestResult.FAILURE:
                return "FAIL";
            case ITestResult.SKIP:
                return "SKIP";
            case ITestResult.SUCCESS_PERCENTAGE_FAILURE:
                return "PERCENTAGE FAIL";
            case ITestResult.STARTED:
                return "SKIP";
            default:
                throw new AssertionError("Unexpected value: " + num);
        }
    }

    public Map<String, List<String>> getExcelValue(String filePath) {
        Map<String, List<String>> map = new TreeMap<>();
        try {
            FileInputStream fileInputStream = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheetAt(1);
            Row heardRow = sheet.getRow(0);
            int caseIDIndex = 0;
            for (int i = 0; i < heardRow.getPhysicalNumberOfCells(); i++) {
                String caseID = heardRow.getCell(i).getStringCellValue();
                if (caseID.equals("ID"))
                    caseIDIndex = i;
                break;
            }
            for (Row row : sheet) {
                if (row.getRowNum() == 0)
                    continue;
                if (row.getCell(caseIDIndex) != null) {
                    List<String> caseValues = new ArrayList<>();
                    String key = row.getCell(caseIDIndex).getStringCellValue().replace(" ", "");
                    System.out.println("key:" + key);
                    for (Cell cell : row) {
                        if (cell.getColumnIndex() == caseIDIndex) {
                            continue;
                        } else {
//                            System.out.println("cell.value:"+cell.getStringCellValue());
//                            caseValues.add(cell.getStringCellValue());
                            caseValues.add(getCellValue(cell));
                        }
                    }
                    map.put(key, caseValues);
                } else
                    break;


            }

//            for (int i = 1; i < 20; i++) {
//                row=sheet.getRow(i);
//                List<String> caseRow = new ArrayList<>();
//                for (int j =1; j < row.getPhysicalNumberOfCells(); j++) {
//                    caseRow.add(row.getCell(j).getStringCellValue());
//                }
//                String caseID = row.getCell(0).getStringCellValue().replace(" ","");
//                map.put(caseID, caseRow);
//            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return map;
    }


    public static String getCellValue(Cell cell) {
        String value;
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue() + "";
                } else {
                    value = cell.getNumericCellValue() + "";
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue() + "";
                break;
            default:
                value = cell.getStringCellValue();
        }
        return value;
    }

    public List<String> getExcelHeardRow(String filePath) {

        List<String> heardList = new ArrayList<>();
        try {
            heardList.add("No.");
            File file = new File(filePath);
            FileInputStream fileInputStream = new FileInputStream(file);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheetAt(1);
            Row heardRow = sheet.getRow(0);
            for (int i = 0; i < heardRow.getPhysicalNumberOfCells(); i++) {
                heardList.add(heardRow.getCell(i).getStringCellValue());
            }
            heardList.add("Results");
            heardList.add("Comments");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return heardList;
    }

    public void writeExcel(String sheetName, String filePath, Map<String, List<String>> testCaseMap) {
        try {
            SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(200);
            Sheet sheet = sxssfWorkbook.createSheet(sheetName);

            sheet.setDefaultColumnWidth(30);

            CellStyle heardStyle=sxssfWorkbook.createCellStyle();
            Font heardFont=sxssfWorkbook.createFont(); //设置字体
            heardFont.setFontHeightInPoints((short)12);
            heardFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
            heardStyle.setFont(heardFont);
            heardStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); //设置居中
            heardStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); //
            heardStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());//前景颜色
            heardStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);//填充方式，前涩填充
            heardStyle.setWrapText(true); //自动换行
            heardStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);//下边框
            heardStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
            heardStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
            heardStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
            heardStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            heardStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            heardStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            heardStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());


            CellStyle textStyle=sxssfWorkbook.createCellStyle();
            Font textFont=sxssfWorkbook.createFont(); //设置字体
            textFont.setFontHeightInPoints((short)11);
            textStyle.setFont(textFont);
            textStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            textStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
            textStyle.setWrapText(true);
            textStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
            textStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
            textStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
            textStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            textStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            textStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());

            CellStyle passStyle=sxssfWorkbook.createCellStyle();
            Font passFont=sxssfWorkbook.createFont(); //设置字体
            passFont.setFontHeightInPoints((short)11);
            passFont.setColor(IndexedColors.GREEN.getIndex());
            passStyle.setFont(passFont);
            passStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            passStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
            passStyle.setWrapText(true);
            passStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
            passStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
            passStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
            passStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            passStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            passStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());


            CellStyle failStyle=sxssfWorkbook.createCellStyle();
            Font failFont=sxssfWorkbook.createFont(); //设置字体
            failFont.setFontHeightInPoints((short)11);
            failFont.setColor(IndexedColors.RED.getIndex());
            failStyle.setFont(failFont);
            failStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            failStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
            failStyle.setWrapText(true);
            failStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
            failStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
            failStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
            failStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            failStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            failStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());

            Row row = sheet.createRow(0);
            row.setHeightInPoints(30);
            List<String> heardList = getExcelHeardRow(filePath);
            for (int i = 0; i < heardList.size(); i++) {
                Cell heardCell= row.createCell(i);
                heardCell.setCellValue(heardList.get(i));
                heardCell.setCellStyle(heardStyle);

            }
            Map<String, List<String>> map = testCaseMap;
            List<List<String>> allValues = new ArrayList<>(map.values());
            for (int k = 0; k < map.size(); k++) {
                Row caseRow = sheet.createRow(k + 1);
                caseRow.setHeightInPoints(30);
                Cell firstCell=caseRow.createCell(0);
                firstCell.setCellValue(k + 1);
                firstCell.setCellStyle(textStyle);
                List<String> caseConnet = allValues.get(k);
                for (int j = 0; j < caseConnet.size(); j++) {
                    Cell textCell= caseRow.createCell(j + 1);
                    String result=caseConnet.get(j);
                    if(result.equals("PASS")||result.equals("SKIP")){
                        textCell.setCellStyle(passStyle);
                    }else if(result.equals("FAIL")){
                        textCell.setCellStyle(failStyle);
                    }else{
                        textCell.setCellStyle(textStyle);
                    }
                    textCell.setCellValue(caseConnet.get(j));


                }
            }

            String fileP = "E:" + FILE_SEPARATOR + "Android_O_AUTO" + FILE_SEPARATOR + "MyAutoTest" + FILE_SEPARATOR +
                    "TestOutReport1";
            File file1 = new File(fileP);
            if (!file1.exists())
                file1.mkdirs();
            String file = fileP + FILE_SEPARATOR + sheetName + ".xlsx";
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            sxssfWorkbook.write(fileOutputStream);
            fileOutputStream.flush();
            fileOutputStream.close();


            String moduleDir = System.getProperty("user.dir");
            System.out.println("currentDir="+moduleDir);
            if(moduleDir.contains("MyAutoTest")){
                File moduleFile = new File(moduleDir);
                moduleDir = moduleFile.getParent();
            }
            System.out.println("afterCurrentDir="+moduleDir);
            String mo=moduleDir.split("\\\\").toString();
            System.out.println("mo="+mo);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
