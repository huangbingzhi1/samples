package com.hisense.util;

import com.hisense.entity.Enterprise;
import com.hisense.entity.Goods;
import com.hisense.entity.People;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * @Author Huang.bingzhi
 * @Date 2019/6/12 10:13
 * @Version 1.0
 * Excel
 */
public class ExcelFactory {
    private final String SHEET_ENTERPRISE = "商家与人员对应关系明细";
    private final String SHEET_PEOPLE_SAMPLE = "人员样本数量";
    private final String SHEET_EXCEPT_SAMPLE = "已调研明细";
    /**
     * 商家Sheet页数据起始行
     */
    private final int LINE_START_ENTERPRISE = 3;
    /**
     * 人员Sheet页数据起始行
     **/
    private final int LINE_START_PEOPLE_AMPLE = 1;
    /**
     * 已用Sheet页数据起始行
     */
    private final int LINE_START_EXCEPT_AMPLE = 3;
    /**
     * 文件后缀名
     */
    private static final String FILE_SUFFIX = ".xlsx";
    /**
     * 被替换的文件名后半部分（假如文件名为abc.xlsx，那处理后的文件名为abc-result.xlsx）
     */
    private static final String RESULT_FILE_SUFFIX = "-result.xlsx";
    private CellStyle normalCellStyle;
    private CellStyle titleCellStyle;

    private String filePath;
    private FileInputStream inputStream;
    private Workbook workbook;
    private Set<String> exceptSampleSet;
    /**
     * 如果cis对应了多条数据，则保存在这里面
     */
    private Map<String, ArrayList<Enterprise>> duplicateEnterpriseMap;
    private Map<String, Enterprise> enterpriseMap;
    private Map<String, People> peopleMap;

    /**
     * 存放人员的结果
     */
    private Map<String, Integer> rPeopleSampleMap = new HashMap<>();
    /**
     * 存放商家的结果
     */
    private Map<String, HashSet<String>> rEnterprisePeopleMap = new HashMap<>();
    /**
     * 存放已调研的人员数量
     */
    private Map<String, Integer> rPeopleSuccess = new HashMap<>();
    /**
     * 用于存放过人员的过程数据
     */
    private Map<String, Integer> pPeopleSampleMap = new HashMap<>();
    /**
     * 存放商家的过程数据
     */
    private Map<String, HashSet<String>> pEnterprisePeopleMap = new HashMap<>();

    /**
     * 商家表中存在，但是并不在人员表中
     */
    private Set<String> noExistsaler;

    public static void main(String[] args) throws Exception {
        if (args.length <= -1) {
            System.out.println("输入有误");
        } else {
//            String filePath = args[0];
            String filePath = "E:\\aaa.xlsx";
            ExcelFactory factory = new ExcelFactory(filePath);

            try {
                System.out.println("正在处理文件：" + filePath);
                factory.doBusiness();
                System.out.println("***************处理成功***************");
            } catch (Exception e) {
                System.out.println(e.toString());
            }

        }
    }

    public ExcelFactory(String filePath) throws Exception {
        this.filePath = filePath;
        inputStream = new FileInputStream(new File(filePath));
        workbook = StreamingReader.builder()
                .rowCacheSize(100)
                .bufferSize(4096)
                .open(inputStream);
        enterpriseMap = new HashMap<>();
        duplicateEnterpriseMap = new HashMap<>();
        peopleMap = new HashMap<>();
        noExistsaler = new HashSet<>();
        exceptSampleSet = new HashSet<>();
    }

    /**
     * 从商家sheet页中读取商家初始数据
     */
    private void readInitEnterpriseInfo() {
        Sheet sheet = workbook.getSheet(SHEET_ENTERPRISE);
        //遍历所有的行
        int i = 0;
        for (Row row : sheet) {
            HashSet<String> rowPeoples = new HashSet<>();
            //数据从第三行开始（0代表第一行）
            if (i++ < LINE_START_ENTERPRISE) {
                continue;
            }
            String cis = getValue(row.getCell(3));
            //忽略已抽样的cis店家
            if ("" == cis || exceptSampleSet.contains(cis)) {
                continue;
            }
            Map<String, People> inspector = new HashMap<>(6);
            Map<String, People> saler = new HashMap<>(6);
            inspector.put(Goods.HXDS.name(), new People(getValue(row.getCell(16)), getValue(row.getCell(17))));
            inspector.put(Goods.HXKT.name(), new People(getValue(row.getCell(18)), getValue(row.getCell(19))));
            inspector.put(Goods.KLKT.name(), new People(getValue(row.getCell(20)), getValue(row.getCell(21))));
            inspector.put(Goods.HXBL.name(), new People(getValue(row.getCell(22)), getValue(row.getCell(23))));
            inspector.put(Goods.RSBL.name(), new People(getValue(row.getCell(24)), getValue(row.getCell(25))));
            inspector.put(Goods.XYJ.name(), new People(getValue(row.getCell(26)), getValue(row.getCell(27))));

            saler.put(Goods.HXDS.name(), new People(getValue(row.getCell(28)), getValue(row.getCell(29))));
            saler.put(Goods.HXKT.name(), new People(getValue(row.getCell(30)), getValue(row.getCell(31))));
            saler.put(Goods.KLKT.name(), new People(getValue(row.getCell(32)), getValue(row.getCell(33))));
            saler.put(Goods.HXBL.name(), new People(getValue(row.getCell(34)), getValue(row.getCell(35))));
            saler.put(Goods.RSBL.name(), new People(getValue(row.getCell(36)), getValue(row.getCell(37))));
            saler.put(Goods.XYJ.name(), new People(getValue(row.getCell(38)), getValue(row.getCell(39))));
            //遍历该行所有产品/客户经理编码
            for (People p : saler.values()) {
                if (!rPeopleSampleMap.containsKey(p.getPCode())) {
                    noExistsaler.add(p.getPCode());
                } else if (pPeopleSampleMap.getOrDefault(p.getPCode(), 0) > 0) {
                    rowPeoples.add(p.getPCode());
                }
            }
            Enterprise enterprise = new Enterprise(getValue(row.getCell(0)),
                    getValue(row.getCell(1)), getValue(row.getCell(2)),
                    getValue(row.getCell(3)), getValue(row.getCell(4)),
                    getValue(row.getCell(5)), getValue(row.getCell(6)),
                    getValue(row.getCell(7)), getValue(row.getCell(8)),
                    getValue(row.getCell(9)), getValue(row.getCell(10)),
                    getValue(row.getCell(11)), getValue(row.getCell(12)),
                    getValue(row.getCell(13)), getValue(row.getCell(14)),
                    getValue(row.getCell(15)), inspector, saler);
            //往商家过程Map里添加商家-人员信息
            if (pEnterprisePeopleMap.containsKey(enterprise.getCisCode())) {
                HashSet<String> peoples = pEnterprisePeopleMap.get(enterprise.getCisCode());
                if (null != peoples) {
                    peoples.addAll(rowPeoples);
                } else {
                    pEnterprisePeopleMap.put(enterprise.getCisCode(), rowPeoples);
                }
            } else {
                pEnterprisePeopleMap.put(enterprise.getCisCode(), rowPeoples);
            }
            //如果该CIS存在多行数据，则把数据存到duplicateEnterpriseMap中
            if (enterpriseMap.containsKey(enterprise.getCisCode())) {
                if (duplicateEnterpriseMap.containsKey(enterprise.getCisCode())) {
                    duplicateEnterpriseMap.get(enterprise.getCisCode()).add(enterprise);
                } else {
                    ArrayList<Enterprise> enterprises = new ArrayList<>();
                    //第一次添加要把enterpriseMap的添加进来
                    enterprises.add(enterpriseMap.get(enterprise.getCisCode()));
                    enterprises.add(enterprise);
                    duplicateEnterpriseMap.put(enterprise.getCisCode(), enterprises);
                }

            } else {
                enterpriseMap.put(enterprise.getCisCode(), enterprise);
            }

        }
    }

    /**
     * 从EXCEL中读取人员初始化信息
     */
    private void readInitPeopleInfo() {
        try {
            Sheet sheet = workbook.getSheet(SHEET_PEOPLE_SAMPLE);
            //遍历所有的行
            int i = 0;
            for (Row row : sheet) {
                //数据从第1行开始（0代表第一行）
                if (i++ < LINE_START_PEOPLE_AMPLE) {
                    continue;
                }
                People people = new People(getValue(row.getCell(0)), getValue(row.getCell(2)), getValue(row.getCell(1)),
                        getValue(row.getCell(3)), getValue(row.getCell(4)), Integer.parseInt(getValue(row.getCell(5))), Integer.parseInt(getValue(row.getCell(6))));
                peopleMap.put(people.getPCode(), people);
                if (people.getNeedSample() > 0) {
                    pPeopleSampleMap.put(people.getPCode(), people.getNeedSample());
                }
            }
            rPeopleSampleMap.putAll(pPeopleSampleMap);
        } catch (Exception e) {
            System.out.println("readInitPeopleInfo-" + e.toString());
        }
    }

    /**
     * 执行操作
     */
    public void doBusiness() {
        //读取人员数据
        readInitPeopleInfo();
        //读取已抽样信息
        exceptSampledInfo();
        //读取商家信息
        readInitEnterpriseInfo();
        try {
            workbook.close();
            inputStream.close();
        } catch (Exception e) {
            System.out.println("doBusiness-" + e.toString());
        }
        extractSample();
        writeResult();
    }

    private void exceptSampledInfo() {
        Sheet sheet = workbook.getSheet(SHEET_EXCEPT_SAMPLE);
        //遍历所有的行
        int j = 0;
        for (Row row : sheet) {
            //数据从第3行开始（0代表第一行）
            if (j++ < LINE_START_EXCEPT_AMPLE) {
                continue;
            }
            String cisStr = getValue(row.getCell(3));
            if (cisStr != "") {
                exceptSampleSet.add(cisStr);
            }
            int peopleColumnStart = 28;
            int peopleColumnEnd = 39;
            for (int k = peopleColumnStart; k < peopleColumnEnd; k = k + 2) {
                String pCodeStr = getValue(row.getCell(k));
                if ("" != pCodeStr) {
                    if (pPeopleSampleMap.containsKey(pCodeStr)) {
                        if (rPeopleSuccess.containsKey(pCodeStr)) {
                            rPeopleSuccess.put(pCodeStr, rPeopleSuccess.get(pCodeStr) + 1);
                        } else {
                            rPeopleSuccess.put(pCodeStr, 1);
                        }
                        Integer old = pPeopleSampleMap.get(pCodeStr);
                        old = old - 3;
                        if (old <= 0) {
                            pPeopleSampleMap.remove(pCodeStr);
                            rPeopleSampleMap.put(pCodeStr, 0);
                        } else {
                            pPeopleSampleMap.put(pCodeStr, old);
                            rPeopleSampleMap.put(pCodeStr, old);
                        }
                    }
                }
            }
        }
    }

    /**
     * 提取样本
     */
    private void extractSample() {
        //记录上一次筛选后，人员样本已经足够的人员编号，此次筛选时，先将其删除
        HashSet<String> exceptPeopleSet = new HashSet<>();
        //循环多轮筛选。每次筛选出一个
        while (pEnterprisePeopleMap.size() > 0 && !endExtractSample()) {
            //a.遍历商家，给商家所拥有的人员统计总样本数
            Iterator<Map.Entry<String, HashSet<String>>> it = pEnterprisePeopleMap.entrySet().iterator();
            int currentNum = 0;
            Map.Entry<String, HashSet<String>> currentEntity = null;
            while (it.hasNext()) {
                Map.Entry<String, HashSet<String>> entity = it.next();
                entity.getValue().removeAll(exceptPeopleSet);
                if (entity.getValue().isEmpty()) {
                    it.remove();
                } else if (entity.getValue().size() > currentNum) {
                    currentEntity = entity;
                    currentNum = entity.getValue().size();
                }
            }
            if (null == currentEntity) {
                break;
            }
            //应该遍历当前的人，根pPeopleSampleMap比较，如果等于0了，则需要把pEnterprisePeopleMap中相关的人删掉。
            Iterator<String> iPeoples = currentEntity.getValue().iterator();
            exceptPeopleSet.clear();
            while (iPeoples.hasNext()) {
                String pCode = iPeoples.next();
                int tempCount = pPeopleSampleMap.get(pCode) - 1;

                //已经完成样本了，此人不需要作为样本了
                if (tempCount < 0) {
                    iPeoples.remove();
                } else {
                    //这个人是最后一个样本
                    if (tempCount == 0) {
                        exceptPeopleSet.add(pCode);
                    }
                    //更新筛选后的人员样本数量
                    pPeopleSampleMap.put(pCode, tempCount);
                }
            }
            //结果商家+1
            rEnterprisePeopleMap.put(currentEntity.getKey(), currentEntity.getValue());
            //过程商家-1
            pEnterprisePeopleMap.remove(currentEntity.getKey());
        }
        //人员结果
        Iterator<Map.Entry<String, Integer>> rPeopleIter = rPeopleSampleMap.entrySet().iterator();
        while (rPeopleIter.hasNext()) {
            Map.Entry<String, Integer> next = rPeopleIter.next();
            try {
                if (pPeopleSampleMap.containsKey(next.getKey())) {
                    next.setValue(next.getValue() - pPeopleSampleMap.get(next.getKey()));
                }
            } catch (Exception e) {
                System.out.println("111");
            }
        }
    }

    private void writeResult() {
        try {
            FileOutputStream outputStream = new FileOutputStream(filePath.replace(FILE_SUFFIX, RESULT_FILE_SUFFIX));
            workbook = new XSSFWorkbook();
            createCellStyle();
            Sheet sheet1 = workbook.createSheet(SHEET_ENTERPRISE);
            writeTitleEnterprise(sheet1);
            writeEnterprise(sheet1);
            Sheet sheet2 = workbook.createSheet(SHEET_PEOPLE_SAMPLE);
            writeTitlePeopleSample(sheet2);
            writePeopleSample(sheet2);
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 生成样式
     */
    private void createCellStyle() {
        normalCellStyle = workbook.createCellStyle();
        normalCellStyle.setBorderBottom(BorderStyle.THIN);
        normalCellStyle.setBorderLeft(BorderStyle.THIN);
        normalCellStyle.setBorderRight(BorderStyle.THIN);
        normalCellStyle.setBorderTop(BorderStyle.THIN);

        titleCellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        titleCellStyle.setFont(font);
        titleCellStyle.setBorderBottom(BorderStyle.THIN);
        titleCellStyle.setBorderLeft(BorderStyle.THIN);
        titleCellStyle.setBorderRight(BorderStyle.THIN);
        titleCellStyle.setBorderTop(BorderStyle.THIN);
        titleCellStyle.setFillForegroundColor(IndexedColors.GOLD.getIndex());
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
    }

    /**
     * 写入人员样本sheet页
     *
     * @param sheet
     */
    private void writePeopleSample(Sheet sheet) {
        int rowIndex = 1;
        for (String pCode : rPeopleSampleMap.keySet()) {
            People p = peopleMap.get(pCode);
            Row newRow = sheet.createRow(rowIndex);
            for (int i = 0; i < 9; i++) {
                Cell cell = newRow.createCell(i);
                cell.setCellStyle(normalCellStyle);
            }
            rowIndex++;
            newRow.getCell(0).setCellValue(p.getCompany());
            newRow.getCell(1).setCellValue(p.getPName());
            newRow.getCell(2).setCellValue(p.getPCode());
            newRow.getCell(3).setCellValue(p.getSecondPart());
            newRow.getCell(4).setCellValue(p.getPosition());
            newRow.getCell(5).setCellValue(p.getNeedSample());
            newRow.getCell(6).setCellValue(rPeopleSuccess.getOrDefault(pCode, 0));
            newRow.getCell(7).setCellValue(p.getNeedSample() - rPeopleSuccess.getOrDefault(pCode, 0) * 3);
            newRow.getCell(8).setCellValue(rPeopleSampleMap.get(pCode));
        }
    }

    /**
     * 写入人员样本sheet页标题
     *
     * @param sheet
     */
    private void writeTitlePeopleSample(Sheet sheet) {
        Row newRow = sheet.createRow(0);
        sheet.setColumnWidth(0, 20 * 256);
        sheet.setColumnWidth(1, 20 * 256);
        sheet.setColumnWidth(2, 20 * 256);
        sheet.setColumnWidth(3, 20 * 256);
        sheet.setColumnWidth(4, 20 * 256);
        sheet.setColumnWidth(5, 20 * 256);
        sheet.setColumnWidth(6, 20 * 256);
        sheet.setColumnWidth(7, 20 * 256);
        sheet.setColumnWidth(8, 20 * 256);
        for (int i = 0; i < 9; i++) {
            Cell cell = newRow.createCell(i);
            cell.setCellStyle(titleCellStyle);
        }
        newRow.getCell(0).setCellValue("分公司");
        newRow.getCell(1).setCellValue("姓名");
        newRow.getCell(2).setCellValue("员工编码");
        newRow.getCell(3).setCellValue("二级部门");
        newRow.getCell(4).setCellValue("职位名称");
        newRow.getCell(5).setCellValue("总共需要样本数");
        newRow.getCell(6).setCellValue("已调查成功数");
        newRow.getCell(7).setCellValue("还需样本数");
        newRow.getCell(8).setCellValue("本次提取样本数");
    }

    /**
     * 写入商家sheet页
     *
     * @param sheet
     */
    private void writeEnterprise(Sheet sheet) {
        int rowIndex = 0;
        Row newRow = null;
        rowIndex++;
        rowIndex = LINE_START_ENTERPRISE;
        for (String cis : rEnterprisePeopleMap.keySet()) {
            if (duplicateEnterpriseMap.containsKey(cis)) {
                ArrayList<Enterprise> enterprises = duplicateEnterpriseMap.get(cis);
                for (Enterprise e : enterprises) {
                    newRow = sheet.createRow(rowIndex);
                    rowIndex++;
                    writeOneEnterprise(rowIndex, cis, newRow, e);
                }

            } else {
                newRow = sheet.createRow(rowIndex);
                rowIndex++;
                Enterprise enterprise = enterpriseMap.get(cis);
                writeOneEnterprise(rowIndex, cis, newRow, enterprise);

            }
        }

    }

    /**
     * 写入商家sheet页标题
     *
     * @param sheet
     */
    private void writeTitleEnterprise(Sheet sheet) {
        sheet.setColumnWidth(2, 20 * 256);
        sheet.setColumnWidth(9, 20 * 256);
        sheet.setColumnWidth(10, 20 * 256);
        sheet.setColumnWidth(13, 20 * 256);
        Row newRow;
        Cell cell;
        int columnIndex = 0;
        for (int j = 0; j < LINE_START_ENTERPRISE; j++) {
            newRow = sheet.createRow(j);
            for (int i = 0; i < 40; i++) {
                cell = newRow.createCell(i);
                cell.setCellStyle(titleCellStyle);
            }
        }
        for (int i = 0; i < 14; i++) {
            CellRangeAddress region = new CellRangeAddress(0, 2, i, i);
            sheet.addMergedRegion(region);
        }
        CellRangeAddress region = new CellRangeAddress(0, 1, 14, 15);
        sheet.addMergedRegion(region);
        region = new CellRangeAddress(0, 0, 16, 27);
        sheet.addMergedRegion(region);
        region = new CellRangeAddress(0, 0, 28, 39);
        sheet.addMergedRegion(region);
        for (int i = 16; i < 27; i = i + 2) {
            region = new CellRangeAddress(1, 1, i, i + 1);
            sheet.addMergedRegion(region);
        }
        for (int i = 28; i < 39; i = i + 2) {
            region = new CellRangeAddress(1, 1, i, i + 1);
            sheet.addMergedRegion(region);
        }
        Row row1 = sheet.getRow(0);
        row1.getCell(0).setCellValue("编码");
        row1.getCell(1).setCellValue("分公司");
        row1.getCell(2).setCellValue("商家全称");
        row1.getCell(3).setCellValue("CIS编码");
        row1.getCell(4).setCellValue("省");
        row1.getCell(5).setCellValue("市");
        row1.getCell(6).setCellValue("区县");
        row1.getCell(7).setCellValue("城市级别");
        row1.getCell(8).setCellValue("营销模式");
        row1.getCell(9).setCellValue("商家经理联系人");
        row1.getCell(10).setCellValue("商家经理联系人电话");
        row1.getCell(11).setCellValue("门店名称");
        row1.getCell(12).setCellValue("门店编码");
        row1.getCell(13).setCellValue("办事处名称");
        row1.getCell(14).setCellValue("办事处经理");
        row1.getCell(16).setCellValue("产品总监");
        row1.getCell(28).setCellValue("产品/客户经理");
        Row row2 = sheet.getRow(1);
        row2.getCell(14).setCellValue("编码");
        row2.getCell(15).setCellValue("姓名");

        columnIndex = 16;
        for (Goods g : Goods.values()) {
            row2.getCell(columnIndex).setCellValue(g.getLabel());
            columnIndex += 2;
        }
        columnIndex = 28;
        for (Goods g : Goods.values()) {
            row2.getCell(columnIndex).setCellValue(g.getLabel());
            columnIndex += 2;
        }
        Row row3 = sheet.getRow(2);
        for (int i = 16; i < 40; i++) {
            if (i % 2 == 0) {
                row3.getCell(i).setCellValue("编码");
            } else {
                row3.getCell(i).setCellValue("姓名");
            }
        }
    }

    private void writeOneEnterprise(int rowIndex, String cis, Row newRow, Enterprise enterprise) {
        for (int i = 0; i < 40; i++) {
            Cell cell = newRow.createCell(i);
            cell.setCellStyle(normalCellStyle);
        }
        if (null != enterprise) {
            newRow.getCell(0).setCellValue(rowIndex - LINE_START_ENTERPRISE);
            newRow.getCell(1).setCellValue(enterprise.getCompany());
            newRow.getCell(2).setCellValue(enterprise.getEName());
            newRow.getCell(3).setCellValue(enterprise.getCisCode());
            newRow.getCell(4).setCellValue(enterprise.getProvince());
            newRow.getCell(5).setCellValue(enterprise.getCity());
            newRow.getCell(6).setCellValue(enterprise.getDistrict());
            newRow.getCell(7).setCellValue(enterprise.getMarketLevel());
            newRow.getCell(8).setCellValue(enterprise.getMarketModel());
            newRow.getCell(9).setCellValue(enterprise.getManager());
            newRow.getCell(10).setCellValue(enterprise.getManagerPhone());
            newRow.getCell(11).setCellValue(enterprise.getStoreName());
            newRow.getCell(12).setCellValue(enterprise.getStoreCode());
            newRow.getCell(13).setCellValue(enterprise.getOfficeName());
            newRow.getCell(14).setCellValue(enterprise.getOfficeManagerCode());
            int columnIndex = 16;
            for (Goods g : Goods.values()) {
                People people = enterprise.getInspector().get(g.name());
                if (null == people || null == people.getPCode() || people.getPCode().length() == 0) {
                    columnIndex += 2;
                    continue;
                }
                newRow.getCell(columnIndex).setCellValue(people.getPCode());
                newRow.getCell(columnIndex + 1).setCellValue(people.getPName());
                columnIndex += 2;
            }
            columnIndex = 28;
            for (Goods g : Goods.values()) {
                People people = enterprise.getSaler().get(g.name());
                if (null == people || null == people.getPCode() || people.getPCode().length() == 0) {
                    columnIndex += 2;
                    continue;
                }
                if (noExistsaler.contains(people.getPCode())) {
                    newRow.getCell(columnIndex).setCellValue("【失效】" + people.getPCode());
                    newRow.getCell(columnIndex + 1).setCellValue("【失效】" + people.getPName());
                } else {
                    if (rEnterprisePeopleMap.get(cis).contains(people.getPCode())) {
                        newRow.getCell(columnIndex).setCellValue(people.getPCode());
                        newRow.getCell(columnIndex + 1).setCellValue(people.getPName());
                    } else {
                        newRow.getCell(columnIndex).setCellValue("【无需】" + people.getPCode());
                        newRow.getCell(columnIndex + 1).setCellValue("【无需】" + people.getPName());
                    }
                }
                columnIndex += 2;
            }

        }
    }

    /**
     * 是否符合停止提取样本的条件
     * 如果所有人员所需的样本数目为0，则停止，否则继续
     *
     * @return
     */
    private boolean endExtractSample() {
        for (Integer i : pPeopleSampleMap.values()) {
            if (i > 0) {
                return false;
            }
        }
        return true;
    }

    /**
     * 取出每列的值
     *
     * @param xCell 列
     * @return
     */
    private String getValue(Cell xCell) {
        if (xCell == null) {
            return "";
        }
        if (xCell.getCellType() == CellType.BOOLEAN) {
            return String.valueOf(xCell.getBooleanCellValue());
        } else if (xCell.getCellType() == CellType.NUMERIC) {
            return Double.valueOf(xCell.getNumericCellValue()).longValue() + "";
        } else {
            String temp = xCell.getStringCellValue();
            if (temp.equals("ERROR:  #N/A")) {
                return "";
            } else {
                return temp;
            }
        }
    }
}