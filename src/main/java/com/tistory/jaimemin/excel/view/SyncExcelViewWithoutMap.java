package com.tistory.jaimemin.excel.view;

import com.tistory.jaimemin.excel.builder.SxssfExcelBuilder;
import com.tistory.jaimemin.excel.constant.ExcelConstants;
import com.tistory.jaimemin.excel.service.ExampleService;
import com.tistory.jaimemin.excel.vo.ExampleVO;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Slf4j
@RequiredArgsConstructor
@Component("syncExcelViewWithoutMap")
public class SyncExcelViewWithoutMap extends AbstractXlsView {

    private final static int EXCEL_SIZE = 10000;

    private final ExampleService exampleService;

    @Override
    protected void buildExcelDocument(Map<String, Object> model
            , Workbook wb
            , HttpServletRequest request
            , HttpServletResponse response) throws Exception {
        long listSize = exampleService.getCount();

        SXSSFWorkbook sxssfWorkbook = null;

        for (int start = 0; start < listSize; start += EXCEL_SIZE) {
            Page<ExampleVO> page
                    = exampleService.getExcelVOs(PageRequest.of(start / EXCEL_SIZE, EXCEL_SIZE));

            sxssfWorkbook = SxssfExcelBuilder.createExcelWithoutMap(null
                    , page
                    , start
                    , start == 0 ? null : sxssfWorkbook);
        }

        if (listSize == 0) {
            sxssfWorkbook = SxssfExcelBuilder.createExcelWithoutMap(null
                    , exampleService.getExcelVOs(PageRequest.of(0, EXCEL_SIZE))
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
}
