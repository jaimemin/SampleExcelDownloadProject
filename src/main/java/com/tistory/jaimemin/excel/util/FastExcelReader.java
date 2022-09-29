package com.tistory.jaimemin.excel.util;

import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.MapperFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ObjectUtils;
import org.dhatim.fastexcel.reader.Cell;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

@Slf4j
public class FastExcelReader {

    public static <E> List<E> read(InputStream is, Map<String, Type> cells, Class<E> classType) {
        List<E> result = new ArrayList<>();

        if (ObjectUtils.isEmpty(is)) {
            return result;
        }

        try (ReadableWorkbook wb = new ReadableWorkbook(is)) {
            ObjectMapper objectMapper = new ObjectMapper();
            objectMapper.configure(DeserializationFeature.FAIL_ON_IGNORED_PROPERTIES, false);
            objectMapper.configure(MapperFeature.ACCEPT_CASE_INSENSITIVE_PROPERTIES, true);

            wb.getSheets().forEach(sheet -> {
                try (Stream<Row> rows = sheet.openStream()) {
                    /**
                     * 각 sheet당 첫 번째 row는 header
                     */
                    rows.skip(1).forEach(row -> {
                        AtomicInteger index = new AtomicInteger();
                        /**
                         * 각 row마다의 값을 저장할 맵 객체 저장되는 형식은 다음과 같다. put("A", "이름"); put("B", "게임명");
                         */
                        Map<String, Object> rowData = new LinkedHashMap<>();

                        cells.forEach((key, val) -> {
                            Object refVal;

                            if (val.equals(Integer.class)
                                    || val.equals(Double.class)
                                    || val.equals(Long.class)) {
                                BigDecimal bigDecimal = row.getCellAsNumber(index.getAndIncrement()).orElse(null);

                                if (bigDecimal == null) {
                                    refVal = 0;
                                } else {
                                    if (val.equals(Integer.class)) {
                                        refVal = bigDecimal.intValue();
                                    } else if (val.equals(Double.class)) {
                                        refVal = bigDecimal.doubleValue();
                                    } else if (val.equals(Long.class)) {
                                        refVal = bigDecimal.longValue();
                                    } else {
                                        refVal = 0;
                                    }
                                }
                            } else if (val.equals(LocalDateTime.class) || val.equals(Date.class)) {
                                refVal = row.getCellAsDate(index.getAndIncrement()).orElse(null);
                            } else if (val.equals(Boolean.class)) {
                                refVal = row.getCellAsBoolean(index.getAndIncrement()).orElse(null);
                            } else {
                                Cell cell = row.getCell(index.getAndIncrement());
                                refVal = cell.getValue() + "";

                                // [FastExcelReader.read] ERROR Wrong cell type NUMBER, wanted STRING
                                // refVal = row.getCellAsString(index.getAndIncrement()).orElse("-");
                            }

                            rowData.put(key, refVal);
                        });

                        E o = objectMapper.convertValue(rowData, classType);
                        result.add(o);
                    });
                } catch (IOException e) {
                    log.error("[FastExcelReader.read] ERROR {}", e.getMessage());
                }
            });

            return result;
        } catch (Exception e) {
            log.error("[FastExcelReader.read] ERROR {}", e.getMessage());
        }

        return result;
    }
}

