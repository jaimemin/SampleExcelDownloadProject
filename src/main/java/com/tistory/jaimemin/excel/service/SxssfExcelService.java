package com.tistory.jaimemin.excel.service;

import com.tistory.jaimemin.excel.builder.SxssfExcelBuilder;
import com.tistory.jaimemin.excel.constant.ExcelConstants;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Slf4j
public abstract class SxssfExcelService {

    private static final int PAGE_SIZE = 10000;

    public SXSSFWorkbook getWorkbook(Map<String, Object> excelMap) throws IOException {
        List<String> keys = (List<String>)excelMap.get(ExcelConstants.KEYS);
        List<String> headers = (List<String>)excelMap.get(ExcelConstants.HEADERS);
        long listSize = (long) excelMap.get(ExcelConstants.LIST_SIZE);
        SXSSFWorkbook sxssfWorkbook = null;

        try {
            for (int start = 0; start < listSize; start += PAGE_SIZE) {
                List<Map<String, Object>> list = getExcelList(excelMap, start, PAGE_SIZE);

                sxssfWorkbook = SxssfExcelBuilder.createExcel(headers
                        , keys
                        , null
                        , null
                        , list
                        , start
                        , start == 0 ? null : sxssfWorkbook);

                list.clear();
            }
        } catch (Exception e) {
            log.error("[SxssfExcelService] error message: {}", e.getMessage());
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

        return sxssfWorkbook;
    }

    public abstract List<Map<String, Object>> getExcelList(Map<String, Object> excelMap
            , int start
            , int size) throws ParseException;
}
