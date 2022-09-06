package com.tistory.jaimemin.excel.view;

import com.tistory.jaimemin.excel.builder.SxssfExcelBuilder;
import com.tistory.jaimemin.excel.constant.ExcelConstants;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Slf4j
public abstract class SyncExcelView extends AbstractXlsView {

    private final static int EXCEL_SIZE = 10000;

    @Override
    protected void buildExcelDocument(Map<String, Object> model
            , Workbook workbook
            , HttpServletRequest request
            , HttpServletResponse response) throws Exception {
        Map<String, Object> excelMap = (Map<String, Object>) model.get(ExcelConstants.EXCEL_MAP);
        List<String> keys = (List<String>) excelMap.get(ExcelConstants.KEYS);
        List<String> headers = (List<String>) excelMap.get(ExcelConstants.HEADERS);
        List<String> widths = (List<String>) excelMap.get(ExcelConstants.WIDTHS);
        long listSize = (long) excelMap.get(ExcelConstants.LIST_SIZE);

        SXSSFWorkbook sxssfWorkbook = null;

        for (int start = 0; start < listSize; start += EXCEL_SIZE) {
            List<Map<String, Object>> list = getExcelList(excelMap, start, EXCEL_SIZE);

            sxssfWorkbook = SxssfExcelBuilder.createExcel(headers
                    , keys
                    , widths
                    , null
                    , list
                    , start
                    , start == 0 ? null : sxssfWorkbook);

            list.clear();
        }

        if (listSize == 0) {
            sxssfWorkbook = SxssfExcelBuilder.createExcel(headers
                    , keys
                    , null
                    , null
                    , new ArrayList<>()
                    , 0
                    , null);
        }

        String fileName = "APACHE_POI_SYNC_EXCEL.xlsx";

        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition"
                , "attachment;filename=" + fileName);

        ServletOutputStream out = response.getOutputStream();
        out.flush();
        sxssfWorkbook.write(out);
        out.flush();
        out.close();
    }

    public abstract List<Map<String, Object>> getExcelList(Map<String, Object> excelMap
            , int start
            , int size) throws ParseException;
}
