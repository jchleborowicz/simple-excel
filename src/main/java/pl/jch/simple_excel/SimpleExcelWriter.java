package pl.jch.simple_excel;

import pl.jch.simple_excel.impl.SimpleExcelWriterBuilderImpl;

import java.io.IOException;
import java.io.OutputStream;

public interface SimpleExcelWriter {

    static SimpleExcelWriterBuilder builder() {
        return new SimpleExcelWriterBuilderImpl();
    }

    void write(OutputStream outputStream, DataSet... dataSets) throws IOException;


    void writeToFile(String file, DataSet... dataSets) throws IOException;

}
