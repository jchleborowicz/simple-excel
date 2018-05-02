package pl.jch.simple_excel;

import lombok.Value;

import java.util.List;

@Value(staticConstructor = "of")
public class DataSet {

    private final String sheetName;
    private final List<?> data;
}
