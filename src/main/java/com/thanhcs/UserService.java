package com.thanhcs;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

class UserService {
    List<ExportModel> getUsers() {
        List<ExportModel> userExportList = new ArrayList<>();
        try {
            userExportList.add(new ExportModel("A", "Nguyen", new SimpleDateFormat("MM/dd/yyyy").parse("12/5/1993"), 22, "Computer Science"));
            userExportList.add(new ExportModel("B", "McCord", new SimpleDateFormat("MM/dd/yyyy").parse("5/5/2010"), 16, "NA"));
            userExportList.add(new ExportModel("Alice", "Tran", new SimpleDateFormat("MM/dd/yyyy").parse("1/7/1983"), 30, "Biology"));
            userExportList.add(new ExportModel("Peter", "Pan", new SimpleDateFormat("MM/dd/yyyy").parse("2/12/1989"), 27, "Biology"));
        } catch (ParseException e) {
            System.out.println("Construct models failed.");
        }
        return userExportList;
    }
}
