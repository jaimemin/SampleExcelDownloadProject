package com.tistory.jaimemin.excel.view;

import com.monitorjbl.xlsx.StreamingReader;
import com.tistory.jaimemin.excel.builder.SxssfExcelBuilder;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

@Slf4j
@Component("asyncExcelView")
public class AsyncExcelView extends AbstractXlsView {

    @Value("${file.path}")
    private String filePath;

    @Override
    protected void buildExcelDocument(Map<String, Object> model
            , Workbook wb
            , HttpServletRequest request
            , HttpServletResponse response) {
        String fileName = (String) model.get("fileName");
        String totalFilePath = filePath + fileName + ".xlsx";
        File file = new File(totalFilePath);

        if (!file.exists()) {
            log.info("[AsyncExcelView] file not exists");

            return;
        }

        try (
                InputStream is = new FileInputStream(file);
                Workbook workbook = StreamingReader.builder()
                        .rowCacheSize(100)
                        .bufferSize(4096)
                        .open(is);
                SXSSFWorkbook sxssfWorkbook = createSXSSFWorkbook(workbook);
        ) {
            String downloadFileName = file.getName();
            String downloadFileExt = FilenameUtils.getExtension(downloadFileName);
            String[] s = downloadFileName.split("_");
            downloadFileName = String.join("_", Arrays.copyOfRange(s, 0, s.length - 1)) + "." + downloadFileExt;

            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-Disposition"
                    , "attachment;filename=" + downloadFileName);

            if (file.exists()) {
                deleteFile(file);
            }

            try (
                    ServletOutputStream out = response.getOutputStream();
            ) {
                out.flush();
                sxssfWorkbook.write(out);
                out.flush();
            } catch (Exception e) {
                log.error("[AsyncExcelView] ERROR {}", e.getMessage());
            }
        } catch (Exception e) {
            log.error("[AsyncExcelView] ERROR {}", e.getMessage());
        }
    }

    private SXSSFWorkbook createSXSSFWorkbook(Workbook workbook) {
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(-1);
        CellStyle headStyle = SxssfExcelBuilder.makeHeadStyle(sxssfWorkbook);
        CellStyle bodyStyle = SxssfExcelBuilder.makeBodyStyle(sxssfWorkbook);

        for (Sheet sheet : workbook){
            Sheet sxssfSheet = sxssfWorkbook.createSheet(sheet.getSheetName());
            int rowIdx = 0;
            boolean isHeader = true;
            boolean setWidth = false;
            List<Integer> widths = new ArrayList<>();

            for (Row r : sheet) {
                Row sxssfRow = sxssfSheet.createRow(rowIdx++);
                int cellIdx = 0;

                for (Cell c : r) {
                    String cellValue = c.getStringCellValue();

                    if (!setWidth) {
                        widths.add(StringUtils.isEmpty(cellValue)
                                ? 0 : cellValue.length());
                    } else {
                        widths.set(cellIdx
                                , Math.max(StringUtils.isEmpty(cellValue) ? 0 : cellValue.length(), widths.get(cellIdx)));
                    }

                    Cell sxssfCell = sxssfRow.createCell(cellIdx++);
                    sxssfCell.setCellStyle(isHeader ? headStyle : bodyStyle);
                    sxssfCell.setCellValue(cellValue);
                }

                isHeader = false;
                setWidth = true;
            }

            int widthIdx = 0;

            for (int width : widths) {
                sxssfSheet.setColumnWidth(widthIdx++, 256 * Math.max(30, width));
            }
        }

        return sxssfWorkbook;
    }

    private synchronized void deleteFile(File file) {
        file.delete();
    }
}
