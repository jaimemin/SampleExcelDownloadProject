package com.tistory.jaimemin.excel.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.text.ParseException;
import java.util.List;
import java.util.Map;

@Slf4j
@Service
@RequiredArgsConstructor
public class ExampleSxssfExcelService extends SxssfExcelService {

    private final ExampleService exampleService;

    @Override
    public List<Map<String, Object>> getExcelList(Map<String, Object> excelMap
            , int start
            , int size) throws ParseException {
        return exampleService.getListForPoi(start, size);
    }
}
