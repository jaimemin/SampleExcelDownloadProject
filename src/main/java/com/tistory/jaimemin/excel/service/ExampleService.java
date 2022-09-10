package com.tistory.jaimemin.excel.service;

import com.tistory.jaimemin.excel.constant.ExcelConstants;
import com.tistory.jaimemin.excel.repository.ExampleRepository;
import com.tistory.jaimemin.excel.vo.ExampleVO;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Slf4j
@Service
@RequiredArgsConstructor
public class ExampleService {

    private final static int PAGE_SIZE = 10000;

    private final ExampleRepository exampleRepository;

    public Page<ExampleVO> getExcelVOs(PageRequest pageRequest) {
        return exampleRepository.findAll(pageRequest);
    }

    public Long getCount() {
        return exampleRepository.count();
    }

    @PostConstruct
    void init() {
        if (exampleRepository.count() > 0) {
            return;
        }

        List<ExampleVO> exampleVOs = new ArrayList<>();

        for (int i = 0; i < 1000000; i++) {
            ExampleVO exampleVO = ExampleVO.builder()
                    .column1(LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd HH:mm:ss.SSS")))
                    .column2("test")
                    .column3(i)
                    .column4(i * 1L)
                    .column5(0.5)
                    .column6(UUID.randomUUID().toString().substring(0, 8))
                    .column7(UUID.randomUUID().toString().substring(0, 8))
                    .column8(UUID.randomUUID().toString().substring(0, 8))
                    .column9(UUID.randomUUID().toString().substring(0, 8))
                    .column10(UUID.randomUUID().toString().substring(0, 8))
                    .column11(UUID.randomUUID().toString().substring(0, 8))
                    .column12(UUID.randomUUID().toString().substring(0, 8))
                    .column13(UUID.randomUUID().toString().substring(0, 8))
                    .column14(UUID.randomUUID().toString().substring(0, 8))
                    .column15(UUID.randomUUID().toString().substring(0, 8))
                    .column16(UUID.randomUUID().toString().substring(0, 8))
                    .column17(UUID.randomUUID().toString().substring(0, 8))
                    .column18(UUID.randomUUID().toString().substring(0, 8))
                    .column19(UUID.randomUUID().toString().substring(0, 8))
                    .column20(UUID.randomUUID().toString().substring(0, 8))
                    .build();

            exampleVOs.add(exampleVO);

            if (exampleVOs.size() == PAGE_SIZE) {
                exampleRepository.saveAll(exampleVOs);

                exampleVOs.clear();
            }
        }

        exampleRepository.saveAll(exampleVOs);
    }

    public Map<String, Object> getExcelMap() {
        List<String> keys = Arrays.asList("NO"
                , "COLUMN_1"
                , "COLUMN_2"
                , "COLUMN_3"
                , "COLUMN_4"
                , "COLUMN_5"
                , "COLUMN_6"
                , "COLUMN_7"
                , "COLUMN_8"
                , "COLUMN_9"
                , "COLUMN_10"
                , "COLUMN_11"
                , "COLUMN_12"
                , "COLUMN_13"
                , "COLUMN_14"
                , "COLUMN_15"
                , "COLUMN_16"
                , "COLUMN_17"
                , "COLUMN_18"
                , "COLUMN_19"
                , "COLUMN_20");
        List<String> headers = Arrays.asList("No."
                , "첫번 째 칼럼"
                , "두번 째 칼럼"
                , "세번 째 칼럼"
                , "네번 째 칼럼"
                , "다섯번 째 칼럼"
                , "여섯번 째 칼럼"
                , "일곱번 째 칼럼"
                , "여덟번 째 칼럼"
                , "아홉번 째 칼럼"
                , "열번 째 칼럼"
                , "열한번 째 칼럼"
                , "열두번 째 칼럼"
                , "열세번 째 칼럼"
                , "열네번 째 칼럼"
                , "열다섯번 째 칼럼"
                , "열여섯번 째 칼럼"
                , "열일곱번 째 칼럼"
                , "열여덟번 째 칼럼"
                , "열아홉번 째 칼럼"
                , "스무번 째 칼럼");
        List<String> widths = Arrays.asList("50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50"
                , "50");
        List<String> aligns = Arrays.asList("CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER"
                , "CENTER");

        Map<String, Object> map = new HashMap<>();
        map.put(ExcelConstants.HEADERS, headers);
        map.put(ExcelConstants.WIDTHS, widths);
        map.put(ExcelConstants.KEYS, keys);
        map.put(ExcelConstants.ALIGNS, aligns);
        map.put(ExcelConstants.LIST_SIZE, exampleRepository.count());

        return map;
    }

    public List<Map<String, Object>> getListForFastExcel(int idx, int start) {
        Page<ExampleVO> page
                =  getExcelVOs(PageRequest.of(idx / PAGE_SIZE + (start * ExcelConstants.MAX_ROWS_PER_SHEET / PAGE_SIZE), PAGE_SIZE));
        List<Map<String, Object>> list = new ArrayList<>();

        for (ExampleVO exampleVO : page) {
            Map<String, Object> tempMap = new HashMap<>();
            tempMap.put("NO", exampleVO.getId());
            tempMap.put("COLUMN_1", exampleVO.getColumn1());
            tempMap.put("COLUMN_2", exampleVO.getColumn2());
            tempMap.put("COLUMN_3", exampleVO.getColumn3());
            tempMap.put("COLUMN_4", exampleVO.getColumn4());
            tempMap.put("COLUMN_5", exampleVO.getColumn5());
            tempMap.put("COLUMN_6", exampleVO.getColumn6());
            tempMap.put("COLUMN_7", exampleVO.getColumn7());
            tempMap.put("COLUMN_8", exampleVO.getColumn8());
            tempMap.put("COLUMN_9", exampleVO.getColumn9());
            tempMap.put("COLUMN_10", exampleVO.getColumn10());
            tempMap.put("COLUMN_11", exampleVO.getColumn11());
            tempMap.put("COLUMN_12", exampleVO.getColumn12());
            tempMap.put("COLUMN_13", exampleVO.getColumn13());
            tempMap.put("COLUMN_14", exampleVO.getColumn14());
            tempMap.put("COLUMN_15", exampleVO.getColumn15());
            tempMap.put("COLUMN_16", exampleVO.getColumn16());
            tempMap.put("COLUMN_17", exampleVO.getColumn17());
            tempMap.put("COLUMN_18", exampleVO.getColumn18());
            tempMap.put("COLUMN_19", exampleVO.getColumn19());
            tempMap.put("COLUMN_20", exampleVO.getColumn20());

            list.add(tempMap);
        }

        return list;
    }

    public List<Map<String, Object>> getListForPoi(int start, int size) {
        Page<ExampleVO> page = getExcelVOs(PageRequest.of(start / PAGE_SIZE, size));
        List<Map<String, Object>> list = new ArrayList<>();

        for (ExampleVO exampleVO : page) {
            Map<String, Object> tempMap = new HashMap<>();
            tempMap.put("NO", exampleVO.getId());
            tempMap.put("COLUMN_1", exampleVO.getColumn1());
            tempMap.put("COLUMN_2", exampleVO.getColumn2());
            tempMap.put("COLUMN_3", exampleVO.getColumn3());
            tempMap.put("COLUMN_4", exampleVO.getColumn4());
            tempMap.put("COLUMN_5", exampleVO.getColumn5());
            tempMap.put("COLUMN_6", exampleVO.getColumn6());
            tempMap.put("COLUMN_7", exampleVO.getColumn7());
            tempMap.put("COLUMN_8", exampleVO.getColumn8());
            tempMap.put("COLUMN_9", exampleVO.getColumn9());
            tempMap.put("COLUMN_10", exampleVO.getColumn10());
            tempMap.put("COLUMN_11", exampleVO.getColumn11());
            tempMap.put("COLUMN_12", exampleVO.getColumn12());
            tempMap.put("COLUMN_13", exampleVO.getColumn13());
            tempMap.put("COLUMN_14", exampleVO.getColumn14());
            tempMap.put("COLUMN_15", exampleVO.getColumn15());
            tempMap.put("COLUMN_16", exampleVO.getColumn16());
            tempMap.put("COLUMN_17", exampleVO.getColumn17());
            tempMap.put("COLUMN_18", exampleVO.getColumn18());
            tempMap.put("COLUMN_19", exampleVO.getColumn19());
            tempMap.put("COLUMN_20", exampleVO.getColumn20());

            list.add(tempMap);
        }

        return list;
    }
}
