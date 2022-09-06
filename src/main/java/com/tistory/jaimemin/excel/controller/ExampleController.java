package com.tistory.jaimemin.excel.controller;

import com.tistory.jaimemin.excel.constant.ExcelConstants;
import com.tistory.jaimemin.excel.service.ExampleService;
import com.tistory.jaimemin.excel.service.FastExcelService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;

@Slf4j
@Controller
@RequiredArgsConstructor
@RequestMapping("/download")
public class ExampleController {

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
            // fastExcelService.createFastExcelWithoutMap(os);
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

    @GetMapping(value = "/poi/async/downloadExcelView.do", produces = "application/vnd.ms-excel")
    public String downloadAsyncPoiExcel(@RequestParam String fileName, Model model) {
        model.addAttribute("fileName", fileName);

        return "asyncExcelView";
    }

}
