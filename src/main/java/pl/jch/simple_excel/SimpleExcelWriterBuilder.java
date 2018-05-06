package pl.jch.simple_excel;

import java.util.function.BiConsumer;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public interface SimpleExcelWriterBuilder {

    SimpleExcelWriterBuilder workbookType(WorkbookType workbookType);

    SimpleExcelWriterBuilder workbookCustomization(Consumer<Workbook> customizer);

    SimpleExcelWriterBuilder defineStyle(String name, Consumer<CellStyle> initializer);

    SimpleExcelWriterBuilder defineStyle(String name,
                                         BiConsumer<CellStyle, StyleInitializerContext> initializer);

    SimpleExcelWriterBuilder defineFont(String name, Consumer<Font> initializer);

    SimpleExcelWriterBuilder defineFont(String name, String baseFontName, Consumer<Font> initializer);

    SheetBuilder<Object> sheet(String name);

    <T> SheetBuilder<T> sheet(String name, Class<T> rowClass);

    SimpleExcelWriter build();
}
