package com.tistory.jaimemin.excel.service;

import com.tistory.jaimemin.excel.constant.ExcelConstants;
import com.tistory.jaimemin.excel.vo.ExampleVO;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.dhatim.fastexcel.Color;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

@Slf4j
@Service
@RequiredArgsConstructor
public class FastExcelService {

    private final static int PAGE_SIZE = 10000;

    private final static int FLUSH_SIZE = 100;

    private final ExampleService exampleService;

    public void createFastExcel(OutputStream os) throws IOException {
        Workbook wb
                = new Workbook(os, "ExcelApplication", "1.0");

        Map<String, Object> excelMap = exampleService.getExcelMap();
        int count = Integer.valueOf(Math.toIntExact(exampleService.getCount()));
        int forLoopCnt = count / ExcelConstants.MAX_ROWS_PER_SHEET
                + (count % ExcelConstants.MAX_ROWS_PER_SHEET > 0 ? 1 : 0);

        for (int i = 0; i < forLoopCnt; i++) {
            Worksheet ws = wb.newWorksheet(ExcelConstants.SHEET_NAME + (i + 1));

            createWorksheet(ws
                    , excelMap
                    , i
                    , count);
        }

        wb.finish();
    }

    public void createFastExcelWithoutMap(OutputStream os) throws IOException {
        Workbook wb
                = new Workbook(os, "ExcelApplication", "1.0");

        int count = Integer.valueOf(Math.toIntExact(exampleService.getCount()));
        int forLoopCnt = count / ExcelConstants.MAX_ROWS_PER_SHEET
                + (count % ExcelConstants.MAX_ROWS_PER_SHEET > 0 ? 1 : 0);

        for (int i = 0; i < forLoopCnt; i++) {
            Worksheet ws = wb.newWorksheet(ExcelConstants.SHEET_NAME + (i + 1));

            createWorksheetWithoutMap(ws, i, count);
        }

        wb.finish();
    }

    private void createWorksheet(Worksheet ws
            , Map<String, Object> excelMap
            , int start
            , int count) throws IOException {
        List<String> keys = (List<String>) excelMap.get(ExcelConstants.KEYS);
        List<String> headers = (List<String>) excelMap.get(ExcelConstants.HEADERS);
        List<String> widths = (List<String>) excelMap.get(ExcelConstants.WIDTHS);

        for (int i = 0; i < widths.size(); i++) {
            ws.width(i, Integer.valueOf(widths.get(i)));
        }

        for (int i = 0; i < headers.size(); i++) {
            ws.value(0, i, headers.get(i));
        }

        ws.range(0, 0, 0, headers.size() - 1).style().horizontalAlignment("center").set();
        ws.range(0, 0, 0, headers.size() - 1).style().fillColor(Color.LIGHT_GREEN).set();

        int row = 1;

        for (int idx = 0; idx < Math.min(ExcelConstants.MAX_ROWS_PER_SHEET, count - start * ExcelConstants.MAX_ROWS_PER_SHEET); idx += PAGE_SIZE) {
            List<Map<String, Object>> list = exampleService.getListForFastExcel(idx, start);

            for (Map<String, Object> map : list) {
                for (int i = 0; i < keys.size(); i++) {
                    ws.value(row, i, String.valueOf(map.get(keys.get(i))));
                }

                if (++row % FLUSH_SIZE == 0) {
                    ws.flush();
                }
            }
        }

        ws.flush();
        ws.finish();
    }

    private void createWorksheetWithoutMap(Worksheet ws
            , int start
            , int count) throws IOException {
        for (int i = 0; i < 6; i++) {
            ws.width(i, 50);
        }

        ws.value(0, 0, "No.");
        ws.value(0, 1, "첫번 째 칼럼");
        ws.value(0, 2, "두번 째 칼럼");
        ws.value(0, 3, "세번 째 칼럼");
        ws.value(0, 4, "네번 째 칼럼");
        ws.value(0, 5, "다섯번 째 칼럼");

        ws.range(0, 0, 0, 5).style().horizontalAlignment("center").set();
        ws.range(0, 0, 0, 5).style().fillColor(Color.LIGHT_GREEN).set();

        int row = 1;

        for (int idx = 0; idx < Math.min(ExcelConstants.MAX_ROWS_PER_SHEET, count - start * ExcelConstants.MAX_ROWS_PER_SHEET); idx += PAGE_SIZE) {
            Page<ExampleVO> page = exampleService.getExcelVOs(PageRequest.of(idx / PAGE_SIZE + (start * ExcelConstants.MAX_ROWS_PER_SHEET / PAGE_SIZE), PAGE_SIZE));

            for (ExampleVO exampleVO : page) {
                ws.value(row, 0, exampleVO.getId());;
                ws.value(row, 1, exampleVO.getColumn1());
                ws.value(row, 2, exampleVO.getColumn2());
                ws.value(row, 3, exampleVO.getColumn3());
                ws.value(row, 4, exampleVO.getColumn4());
                ws.value(row, 5, exampleVO.getColumn5());

                if (++row % FLUSH_SIZE == 0) {
                    ws.flush();
                }
            }
        }

        ws.flush();
        ws.finish();
    }
}
