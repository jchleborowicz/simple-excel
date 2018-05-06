package pl.jch.simple_excel;

import java.io.IOException;
import java.io.OutputStream;

import pl.jch.simple_excel.impl.SimpleExcelWriterBuilderImpl;

public interface SimpleExcelWriter {

    static SimpleExcelWriterBuilder builder() {
        return new SimpleExcelWriterBuilderImpl();
    }

    void write(OutputStream outputStream, DataSet... dataSets) throws IOException;


    void writeToFile(String file, DataSet... dataSets) throws IOException;

}
