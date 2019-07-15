package com.thanhcs;

import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Date;

@Getter
@AllArgsConstructor
class ExportModel {
    private String firstName;
    private String lastName;
    private Date dob;
    private int age;
    private String major;
}
