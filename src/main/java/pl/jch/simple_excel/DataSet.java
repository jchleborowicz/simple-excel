package pl.jch.simple_excel;

import java.util.List;

import lombok.Value;

@Value(staticConstructor = "of")
public class DataSet {

    private final String sheetName;
    private final List<?> data;
}
