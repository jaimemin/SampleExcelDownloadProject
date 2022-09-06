package com.tistory.jaimemin.excel.builder;

import com.tistory.jaimemin.excel.constant.ExcelConstants;
import com.tistory.jaimemin.excel.vo.ExampleVO;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.data.domain.Page;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.*;

@Slf4j
public class SxssfExcelBuilder {

    private static final String SHEET_NAME = "Sheet";

    private static final String LEFT = "LEFT";

    private static final String CENTER = "CENTER";

    private static final String RIGHT = "RIGHT";

    private static final int MAX_ROW = 1040000;

    private static final int FLUSH_ROW_NUM = 100;

    public static SXSSFWorkbook createExcel(List<String> headers
            , List<String> keys
            , List<String> widths
            , String sheetName
            , List<Map<String,Object>> list
            , int rowIdx
            , SXSSFWorkbook sxssfWorkbook) throws IOException {
        /**
         * 최초 생성이면 manual flush를 위해 new SXSSFWorkbook(-1);
         * 이어서 작성일 경우 매개변수로 받은 sxssfWorkbook
         */
        SXSSFWorkbook workbook = ObjectUtils.isNotEmpty(sxssfWorkbook)
                ? sxssfWorkbook : new SXSSFWorkbook(-1);
        /**
         * 최초 생성이면 SHEET_NAME 시트 생성
         * 이어서 작성일 경우 SHEET_NAME에서 이어서 작성
         */
        sheetName = StringUtils.isEmpty(sheetName)
                ? SHEET_NAME + (rowIdx / MAX_ROW + 1) : sheetName;
        boolean newSheet = ObjectUtils.isEmpty(workbook.getSheet(sheetName));
        Sheet sheet = newSheet ? workbook.createSheet(sheetName) : workbook.getSheet(sheetName);

        Row row = null;
        Cell cell = null;

        // 매개변수로 받은 rowNo부터 이어서 작성
        int rowNo = rowIdx % MAX_ROW;
        int index = 0;

        CellStyle headStyle = makeHeadStyle(workbook);
        CellStyle bodyStyleCenter = makeBodyStyle(workbook, CENTER);
        CellStyle bodyStyleLeft = makeBodyStyle(workbook, LEFT);
        CellStyle bodyStyleRight = makeBodyStyle(workbook, RIGHT);

        /**
         * 셀 내 개행을 위해 추가
         * \r\n 추가 시 개행 가능
         */
        bodyStyleCenter.setWrapText(true);
        bodyStyleLeft.setWrapText(true);
        bodyStyleRight.setWrapText(true);

        if (ObjectUtils.isNotEmpty(widths)) {
            for (String width : widths) {
                sheet.setColumnWidth(index++, Integer.parseInt(width) * 256);
            }
        }

        // 헤더 생성
        if (newSheet) {
            row = sheet.createRow(rowNo);

            index = 0;

            for (String colName : headers) {
                cell = row.createCell(index++);
                cell.setCellStyle(headStyle);
                cell.setCellValue(colName);
            }
        }

        // 데이터와 cell alignment
        for (Map<String,Object> aRow: list) {
            row = sheet.createRow(++rowNo);
            index = 0;

            for (String aKey: keys) {
                if (StringUtils.isEmpty(aKey)) {
                    continue;
                }

                cell = row.createCell(index);
                cell.setCellStyle(bodyStyleCenter);

                Object aValue = aRow.get(aKey);

                if (aValue instanceof BigDecimal) {
                    cell.setCellValue(((BigDecimal) aValue).toString());
                } else if (aValue instanceof Double) {
                    cell.setCellValue(((Double) aValue).toString());
                } else if (aValue instanceof Long) {
                    cell.setCellValue(((Long) aValue).toString());
                } else if (aValue instanceof Integer) {
                    cell.setCellValue(((Integer) aValue).toString());
                } else if (aValue instanceof String[]) {
                    String[] options = (String[]) aValue;
                    DataValidationConstraint constraint = sheet.getDataValidationHelper()
                            .createExplicitListConstraint(options);
                    // firstRow, lastRow, firstCol, lastCol
                    CellRangeAddressList addressList = new CellRangeAddressList(rowNo, rowNo, index, index);
                    DataValidation dataValidation = sheet.getDataValidationHelper().createValidation(constraint, addressList);
                    dataValidation.setSuppressDropDownArrow(true);
                    sheet.addValidationData(dataValidation);

                    cell.setCellValue(ObjectUtils.isNotEmpty(options) ? options[0] : null);
                } else {
                    cell.setCellValue((String) aValue);
                }

                index++;

                // 주기적인 flush 진행
                if (rowNo % FLUSH_ROW_NUM == 0) {
                    ((SXSSFSheet) sheet).flushRows(FLUSH_ROW_NUM);
                }
            }
        }

        return workbook;
    }

    public static SXSSFWorkbook createExcelWithoutMap(String sheetName
            , Page<ExampleVO> page
            , int rowIdx
            , SXSSFWorkbook sxssfWorkbook) throws IOException {
        /**
         * 최초 생성이면 manual flush를 위해 new SXSSFWorkbook(-1);
         * 이어서 작성일 경우 매개변수로 받은 sxssfWorkbook
         */
        SXSSFWorkbook workbook = ObjectUtils.isNotEmpty(sxssfWorkbook)
                ? sxssfWorkbook : new SXSSFWorkbook(-1);
        /**
         * 최초 생성이면 SHEET_NAME 시트 생성
         * 이어서 작성일 경우 SHEET_NAME에서 이어서 작성
         */
        sheetName = StringUtils.isEmpty(sheetName)
                ? SHEET_NAME + (rowIdx / MAX_ROW + 1) : sheetName;
        boolean newSheet = ObjectUtils.isEmpty(workbook.getSheet(sheetName));
        Sheet sheet = newSheet ? workbook.createSheet(sheetName) : workbook.getSheet(sheetName);

        Row row = null;
        Cell cell = null;

        // 매개변수로 받은 rowNo부터 이어서 작성
        int rowNo = rowIdx % MAX_ROW;
        int index = 0;

        CellStyle headStyle = makeHeadStyle(workbook);
        CellStyle bodyStyleCenter = makeBodyStyle(workbook, CENTER);
        CellStyle bodyStyleLeft = makeBodyStyle(workbook, LEFT);
        CellStyle bodyStyleRight = makeBodyStyle(workbook, RIGHT);

        /**
         * 셀 내 개행을 위해 추가
         * \r\n 추가 시 개행 가능
         */
        bodyStyleCenter.setWrapText(true);
        bodyStyleLeft.setWrapText(true);
        bodyStyleRight.setWrapText(true);

        for (int i = 0; i < 6; i++) {
            sheet.setColumnWidth(i, 50 * 256);
        }

        // 헤더 생성
        if (newSheet) {
            row = sheet.createRow(rowNo);

            for (int i = 0; i < 6; i++) {
                cell = row.createCell(i);
                cell.setCellStyle(headStyle);
                cell.setCellValue(getHeader(i));
            }
        }

        // 데이터와 cell alignment
        for (ExampleVO exampleVO : page) {
            row = sheet.createRow(++rowNo);
            index = 0;

            cell = row.createCell(0);
            cell.setCellStyle(bodyStyleCenter);
            cell.setCellValue(exampleVO.getId());

            cell = row.createCell(1);
            cell.setCellStyle(bodyStyleCenter);
            cell.setCellValue(exampleVO.getColumn1());

            cell = row.createCell(2);
            cell.setCellStyle(bodyStyleCenter);
            cell.setCellValue(exampleVO.getColumn2());

            cell = row.createCell(3);
            cell.setCellStyle(bodyStyleCenter);
            cell.setCellValue(exampleVO.getColumn3());

            cell = row.createCell(4);
            cell.setCellStyle(bodyStyleCenter);
            cell.setCellValue(exampleVO.getColumn4());

            cell = row.createCell(5);
            cell.setCellStyle(bodyStyleCenter);
            cell.setCellValue(exampleVO.getColumn5());

            if (rowNo % FLUSH_ROW_NUM == 0) {
                ((SXSSFSheet) sheet).flushRows(FLUSH_ROW_NUM);
            }
        }

        return workbook;
    }

    public static Map<String, Object> makeExcelDataMap(long listSize
            , Object criteria
            , List<String> keys
            , List<String> headers
            , List<String> widths
            , String sheetName) {
        Map<String, Object> map = new HashMap<>();
        map.put(ExcelConstants.HEADERS, headers);
        map.put(ExcelConstants.CRITERIA, criteria);
        map.put(ExcelConstants.KEYS, keys);
        map.put(ExcelConstants.LIST_SIZE, listSize);
        map.put(ExcelConstants.WIDTHS, widths);
        map.put(ExcelConstants.SHEET_NAME, sheetName);

        return map;
    }

    public static CellStyle makeHeadStyle(SXSSFWorkbook workbook) {
        CellStyle headStyle = makeBodyStyle(workbook);
        headStyle.setFillForegroundColor(HSSFColor
                .HSSFColorPredefined
                .PALE_BLUE
                .getIndex());
        headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headStyle.setAlignment(HorizontalAlignment.CENTER);
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        return headStyle;
    }

    public static CellStyle makeBodyStyle(SXSSFWorkbook workbook) {
        CellStyle bodyStyle = workbook.createCellStyle();
        bodyStyle.setBorderTop(BorderStyle.THIN);
        bodyStyle.setBorderBottom(BorderStyle.THIN);
        bodyStyle.setBorderLeft(BorderStyle.THIN);
        bodyStyle.setBorderRight(BorderStyle.THIN);
        bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        bodyStyle.setAlignment(HorizontalAlignment.CENTER);

        return bodyStyle;
    }

    public static CellStyle makeBodyStyle(SXSSFWorkbook workbook, String align) {
        CellStyle bodyStyle = workbook.createCellStyle();
        bodyStyle.setBorderTop(BorderStyle.THIN);
        bodyStyle.setBorderBottom(BorderStyle.THIN);
        bodyStyle.setBorderLeft(BorderStyle.THIN);
        bodyStyle.setBorderRight(BorderStyle.THIN);
        bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        if (StringUtils.isNotEmpty(align)) {
            if (LEFT.equals(align)) {
                bodyStyle.setAlignment(HorizontalAlignment.LEFT);
            } else if (RIGHT.equals(align)) {
                bodyStyle.setAlignment(HorizontalAlignment.RIGHT);
            } else {
                bodyStyle.setAlignment(HorizontalAlignment.CENTER);
            }
        }

        return bodyStyle;
    }

    public static String getHeader(int idx) {
        switch (idx) {
            case 0:
                return "No.";
            case 1:
                return "첫번 째 칼럼";
            case 2:
                return "두번 째 칼럼";
            case 3:
                return "세번 째 칼럼";
            case 4:
                return "네번 째 칼럼";
            case 5:
                return "다섯번 째 칼럼";
            default:
                return null;
        }
    }
}
