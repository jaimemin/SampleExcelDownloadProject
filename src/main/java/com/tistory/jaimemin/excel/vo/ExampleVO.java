package com.tistory.jaimemin.excel.vo;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.Table;

@Data
@Entity
@Builder
@Table(name = "excel")
@NoArgsConstructor
@AllArgsConstructor
public class ExampleVO {

    @Id
    @GeneratedValue
    private Long id;

    private String column1;

    private String column2;

    private Integer column3;

    private Long column4;

    private Double column5;

    private String column6;

}
