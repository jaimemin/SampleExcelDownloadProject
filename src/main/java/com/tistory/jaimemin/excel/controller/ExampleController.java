package com.tistory.jaimemin.excel.controller;

import com.tistory.jaimemin.excel.builder.SxssfExcelBuilder;
import com.tistory.jaimemin.excel.constant.ExcelConstants;
import com.tistory.jaimemin.excel.service.ExampleService;
import com.tistory.jaimemin.excel.service.FastExcelService;
import com.tistory.jaimemin.excel.vo.ExampleVO;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;

@Slf4j
@Controller
@RequiredArgsConstructor
@RequestMapping("/download")
public class ExampleController {

    private final static int EXCEL_SIZE = 10000;

    private final ExampleService exampleService;

    private final FastExcelService fastExcelService;

    @GetMapping("/fastexcel")
    public void downloadExcel(HttpServletResponse response) throws UnsupportedEncodingException {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        String fileNameUtf8 = URLEncoder.encode("FAST_EXCEL", "UTF-8");
        response.setHeader("Content-Disposition", "attachment; filename=" + fileNameUtf8 + ".xlsx");

        try (OutputStream os = response.getOutputStream()) {
            fastExcelService.createFastExcel(os);
        } catch (IOException e) {
            log.error("[fastexcel] ERROR {}", e.getMessage());
        }
    }

    @GetMapping("/fastexcel/mapless")
    public void downloadExcelWithMap(HttpServletResponse response) throws UnsupportedEncodingException {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        String fileNameUtf8 = URLEncoder.encode("FAST_EXCEL", "UTF-8");
        response.setHeader("Content-Disposition", "attachment; filename=" + fileNameUtf8 + ".xlsx");

        try (OutputStream os = response.getOutputStream()) {
            fastExcelService.createFastExcelWithoutMap(os);
        } catch (IOException e) {
            log.error("[fastexcel] ERROR {}", e.getMessage());
        }
    }

    @GetMapping(value = "/poi/sync", produces = "application/vnd.ms-excel")
    public String downloadPoiExcelSyncVersion(Model model) {
        model.addAttribute(ExcelConstants.EXCEL_MAP, exampleService.getExcelMap());

        // return "syncExcelViewWithoutMap";
        return "syncExcelView";
    }

    @GetMapping(value = "/poi/sync/mapless", produces = "application/vnd.ms-excel")
    public String downloadPoiExcelSyncVersionWithMap(Model model) {
        model.addAttribute(ExcelConstants.EXCEL_MAP, exampleService.getExcelMap());

        return "syncExcelViewWithoutMap";
    }

    @GetMapping(value = "/poi/async/downloadExcelView.do", produces = "application/vnd.ms-excel")
    public String downloadAsyncPoiExcel(@RequestParam String fileName, Model model) {
        model.addAttribute("fileName", fileName);

        return "asyncExcelView";
    }

    @GetMapping("/csv")
    public void generateCsv(HttpServletResponse response) {
        String fileName = "testCSV";

        response.addHeader("Content-Type", "application/csv");
        response.addHeader("Content-Disposition", "attachment; filename=" + fileName + ".csv");
        response.setCharacterEncoding("UTF-8");

        try {
            PrintWriter out = new PrintWriter(new OutputStreamWriter(response.getOutputStream(), "UTF8"), true);
            out.write(headerString());
            out.write("\n");
            long listSize = exampleService.getCount();

            for (int start = 0; start < listSize; start += EXCEL_SIZE) {
                Page<ExampleVO> page
                        = exampleService.getExcelVOs(PageRequest.of(start / EXCEL_SIZE, EXCEL_SIZE));

                for (ExampleVO exampleVO : page) {
                    out.write(exampleVO.toString());
                    out.write("\n");
                }

                out.flush();
            }

            out.close();
        } catch (Exception e) {
            log.error("[generateCSV] ERROR {}", e.getMessage());
        }
    }

    private String headerString() {
        return String.join(","
                , "아이디"
                , "첫번 째 칼럼"
                , "두번 째 칼럼"
                , "세번 째 칼럼"
                , "네번 째 칼럼"
                , "다섯번 째 칼럼"
                , "여섯번 째 칼럼");
    }
}
