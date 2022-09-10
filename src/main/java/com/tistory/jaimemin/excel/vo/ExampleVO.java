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

    private String column6;

    private String column7;

    private String column8;

    private String column9;

    private String column10;

    private String column11;

    private String column12;

    private String column13;

    private String column14;

    private String column15;

    private String column16;

    private String column17;

    private String column18;

    private String column19;

    private String column20;

}
