package com.tistory.jaimemin.excel.vo;

import lombok.*;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.Table;
import java.util.Date;

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

}
