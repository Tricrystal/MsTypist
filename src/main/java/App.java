/*
 * Copyright (c) 2021 Tricrystal. All rights reserved.
 */

import java.util.*;
import java.io.*;
import java.io.File;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.apache.poi.xwpf.usermodel.TextSegment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class App {
    private static final String USER_DIR = System.getProperty("user.dir");
    private static final String USER_HOME = System.getProperty("user.home");
    private static final String OUTPUT_BASE_PATH = USER_DIR + "\\output";
    private static final String INPUT_BASE_PATH = USER_DIR + "\\input";

    private static Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        return pattern.matcher(str);
    }


    public static Set<Map<String, String>> fetchExcelContent(File excelFile) throws IOException {
        Set<Map<String, String>> result = new HashSet<>();
        // 工作表对象
        InputStream excelInputStream = new FileInputStream(excelFile);
        XSSFWorkbook workbook = new XSSFWorkbook(excelInputStream);
        excelInputStream.close();
        // 遍历表。
        for (Iterator<Sheet> it = workbook.sheetIterator(); it.hasNext(); ) {
            Sheet sheet = it.next();
            // 行数。
            int rowNumbers = sheet.getLastRowNum() + 1;
            // Excel第一行。
            Row temp = sheet.getRow(0);
            if (temp == null) {
                continue;
            }
            int cellCount = temp.getPhysicalNumberOfCells();
            // 表头
            String[] headers = new String[cellCount];

            Row headerRow = sheet.getRow(0);
            // 填充表头
            for (int col = 0; col < cellCount; col++) {
                headers[col] = headerRow.getCell(col).toString();
            }
//            System.out.println(Arrays.toString(headers));
            // 读数据。
            for (int row = 1; row < rowNumbers; row++) {
                Map<String, String> rowDataMap = new HashMap<>();
                Row r = sheet.getRow(row);
                for (int col = 0; col < cellCount; col++) {
                    rowDataMap.put(headers[col], r.getCell(col) == null ? "" : r.getCell(col).toString());
                }
                result.add(rowDataMap);
            }
        }
        return result;
    }


    public static String replaceWithTag(String oriStr, String target, String replacement) {
        return oriStr.replaceAll( "\\$" + target + "\\$", replacement);
    }

    public static String replaceFromDict(String str, Map<String, String> dict) {
        for (Map.Entry<String, String> entry : dict.entrySet()) {
            str = str.replaceAll("\\$" + entry.getKey() + "\\$", entry.getValue());
        }
        return str;
    }

    /**
     * 单个实体创建对应的文件夹和文件
     *
     * @param baseDocxList baseDocxList docx模板文件
     * @param replaceMap   实体信息 对应Excel里一行
     * @param keyword      实体标注名，用于文件夹名
     * @throws InvalidFormatException POI反序列化异常
     * @throws IOException            IO异常
     */
    public static void copyDocxAndReplace(List<File> baseDocxList, Map<String, String> replaceMap, String keyword) throws
            InvalidFormatException, IOException {
        // 打印键值对
        System.out.println(replaceMap.entrySet().toString());
        // 新建文件夹
        String path = OUTPUT_BASE_PATH + "\\" + replaceMap.get(keyword);
        File pathDir = new File(path); // 建立代表Sub目录的File对象，并得到它的一个引用
        if (!pathDir.exists()) { // 检查Sub目录是否存在
            boolean tag = pathDir.mkdirs();
            if (tag) {
                System.out.println("新建目录： " + path);
            } else {
                System.out.println("新建目录失败： " + path);
            }
        }
        // ?
        for (File baseDocx : baseDocxList) {
            String baseDocxName = baseDocx.getName();
            String newDocxName = replaceFromDict(baseDocxName, replaceMap);
            XWPFDocument document = new XWPFDocument(OPCPackage.open(baseDocx));
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                searchAndReplace(paragraph, replaceMap);
            }
            //  获取表单 遍历表格
            for (XWPFTable table : document.getTables()) {
                // 遍历行
                for (XWPFTableRow tr : table.getRows()) {
                    // 遍历列
                    for (XWPFTableCell tableCell : tr.getTableCells()) {
                        // 遍历单元格中的数据
                        for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                            // 获取单元格中的数据
                            searchAndReplace(paragraph, replaceMap);
                        }
                    }
                }
            }
            FileOutputStream outStream = new FileOutputStream(path + "\\" + newDocxName);
            document.write(outStream);
            outStream.close();
        }
    }

    public static void searchAndReplace(XWPFParagraph paragraph, Map<String, String> map) {
        for (Map.Entry<String, String> key : map.entrySet()) {
            String text = "";
            while (paragraph.getParagraphText().contains("$"+key.getKey()+"$")) {
                PositionInParagraph positionInParagraph = new PositionInParagraph();
                TextSegment textSegement = paragraph.searchText("$"+key.getKey()+"$", positionInParagraph);
                text = paragraph.getText(textSegement).replace("$"+key.getKey()+"$", key.getValue());
                List<XWPFRun> paragraphRuns = paragraph.getRuns();
                for (int i = textSegement.getEndRun(); i > textSegement.getBeginRun(); i--) {
                    paragraph.removeRun(i);
                }
                XWPFRun paragraphRun = paragraphRuns.get(textSegement.getBeginRun());
                CTR ctr = paragraphRun.getCTR();
                for (int i = ctr.sizeOfTArray() - 1; i >= 0; i--) {
                    ctr.removeT(i);
                }
                paragraphRun.setText(text);
            }
        }
    }

    public static List<File> readDocx(String path) {
        if (path.isEmpty()) {
            path = INPUT_BASE_PATH;
        }
        File pathFile = new File(path);
        if (!pathFile.exists()) {
            System.out.println("path=" + pathFile.getPath());
            System.out.println("absolutepath=" + pathFile.getAbsolutePath());
            System.out.println("name=" + pathFile.getName());
            pathFile.mkdirs();
        }
        List<File> fileList = new LinkedList<>();
        String[] fileNameArr = pathFile.list();
        if (fileNameArr != null) {
            for (String fileName : fileNameArr) {
                File file = new File(path + File.separator + fileName);
                if (file.isDirectory() || !fileName.endsWith(".docx")) {
                    continue;
                }
                fileList.add(file);
            }
        }
        return fileList;
    }

    public static List<File> readXlsx(String path) {
        if (path.isEmpty()) {
            path = INPUT_BASE_PATH;
        }
        File pathFile = new File(path);
        if (!pathFile.exists()) {
            System.out.println("path=" + pathFile.getPath());
            System.out.println("absolutepath=" + pathFile.getAbsolutePath());
            System.out.println("name=" + pathFile.getName());
            pathFile.mkdirs();
        }
        List<File> fileList = new LinkedList<>();
        String[] fileNameArr = pathFile.list();
        if (fileNameArr != null) {
            for (String fileName : fileNameArr) {
                File file = new File(path + File.separator + fileName);
                if (file.isDirectory() || !fileName.endsWith(".xlsx")) {
                    continue;
                }
                fileList.add(file);
            }
        }
        return fileList;
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        List<File> docxFileList = readDocx("");
        File excelFile = readXlsx("").get(0);
        Set<Map<String, String>> excelContentList = fetchExcelContent(excelFile);
        for (Map<String, String> excelContentMap : excelContentList) {
            copyDocxAndReplace(docxFileList, excelContentMap, "店铺名");
        }
    }
}
