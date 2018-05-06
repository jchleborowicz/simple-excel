package pl.jch.simple_excel.impl;

import org.apache.commons.lang3.ObjectUtils;
import pl.jch.simple_excel.DataSet;
import pl.jch.simple_excel.SimpleExcelWriter;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

public class SimpleExcelWriterImpl implements SimpleExcelWriter {

    private final SimpleExcelWriterBuilderImpl builder;

    SimpleExcelWriterImpl(SimpleExcelWriterBuilderImpl builder) {
        this.builder = builder;
    }

    @Override
    public void write(OutputStream outputStream, DataSet... dataSets) throws IOException {
        final List<DataSet> dataSetList = Arrays.asList(ObjectUtils.defaultIfNull(dataSets, new DataSet[0]));

        InternalExcelWriter.write(builder, dataSetList, outputStream);
    }

    @Override
    public void writeToFile(String file, DataSet... dataSets) throws IOException {
        try (OutputStream outputStream = new FileOutputStream(file)) {
            write(outputStream, dataSets);
        }
    }

}
