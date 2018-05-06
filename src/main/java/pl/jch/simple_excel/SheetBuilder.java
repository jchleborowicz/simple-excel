package pl.jch.simple_excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.function.BiConsumer;
import java.util.function.Consumer;

public interface SheetBuilder<T> {
    SheetBuilder<Object> sheet(String name);

    <S> SheetBuilder<S> sheet(String name, Class<S> rowClass);

    SheetBuilder<T> sheetCustomization(Consumer<Sheet> customizer);

    ColumnBuilder<T> column(String name);

    SheetBuilder<T> headerStyle(String styleName);

    SheetBuilder<T> headerStyle(Consumer<CellStyle> styleInitializer);

    SheetBuilder<T> headerStyle(String baseStyleName,
                                BiConsumer<CellStyle, StyleInitializerContext> styleInitializer);

    SheetBuilder<T> style(String baseStyleName);

    SheetBuilder<T> style(String baseStyleName,
                          BiConsumer<CellStyle, StyleInitializerContext> styleInitializer);

    SimpleExcelWriter build();
}
