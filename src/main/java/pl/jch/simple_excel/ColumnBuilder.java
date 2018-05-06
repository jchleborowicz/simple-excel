package pl.jch.simple_excel;

import java.util.function.BiConsumer;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;

import org.apache.poi.ss.usermodel.CellStyle;

public interface ColumnBuilder<T> {

    SheetBuilder<Object> sheet(String name);

    <S> SheetBuilder<S> sheet(String name, Class<S> rowClass);

    ColumnBuilder<T> column(String name);

    ColumnBuilder<T> headerStyle(BiConsumer<CellStyle, StyleInitializerContext> styleInitializer);

    ColumnBuilder<T> headerStyle(String baseStyleName,
                                 BiConsumer<CellStyle, StyleInitializerContext> styleInitializer);

    ColumnBuilder<T> style(String baseStyleName);

    ColumnBuilder<T> style(Consumer<CellStyle> styleInitializer);

    ColumnBuilder<T> style(BiConsumer<CellStyle, StyleInitializerContext> styleInitializer);

    ColumnBuilder<T> style(String baseStyleName,
                           BiConsumer<CellStyle, StyleInitializerContext> styleInitializer);

    ColumnBuilder<T> dataExtractor(Function<T, ?> dataValueExtractor);

    ColumnBuilder<T> dataExtractorWithIndex(BiFunction<T, Integer, ?> dataValueExtractor);

    SimpleExcelWriter build();
}
