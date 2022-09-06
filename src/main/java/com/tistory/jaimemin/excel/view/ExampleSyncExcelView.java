package com.tistory.jaimemin.excel.view;


import com.tistory.jaimemin.excel.service.ExampleService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;

@Slf4j
@RequiredArgsConstructor
@Component("syncExcelView")
public class ExampleSyncExcelView extends SyncExcelView {

    private final ExampleService exampleService;

    @Override
    public List<Map<String, Object>> getExcelList(Map<String, Object> excelMap
            , int start
            , int size) {
        return exampleService.getListForPoi(start, size);
    }
}
