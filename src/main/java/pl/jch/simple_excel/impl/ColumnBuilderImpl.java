package pl.jch.simple_excel.impl;

import java.util.function.BiConsumer;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;

import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;
import pl.jch.simple_excel.ColumnBuilder;
import pl.jch.simple_excel.SheetBuilder;
import pl.jch.simple_excel.SimpleExcelException;
import pl.jch.simple_excel.SimpleExcelWriter;
import pl.jch.simple_excel.StyleInitializerContext;

/**
 * @param <T> Row class.
 */
public class ColumnBuilderImpl<T> implements ColumnBuilder<T> {

    private final SheetBuilderImpl<T> parentBuilder;
    @Getter
    final String columnName;
    BiConsumer<CellStyle, StyleInitializerContext> headerStyleInitializer;
    BiConsumer<CellStyle, StyleInitializerContext> columnStyleInitializer;
    BiFunction<T, Integer, ?> dataValueExtractor;


    ColumnBuilderImpl(SheetBuilderImpl<T> parentBuilder, String columnName) {
        this.parentBuilder = parentBuilder;
        this.columnName = columnName;
    }

    @Override
    public SheetBuilder<Object> sheet(String name) {
        return sheet(name, Object.class);
    }

    @Override
    public <S> SheetBuilder<S> sheet(String name, Class<S> rowClass) {
        return this.parentBuilder.sheet(name, rowClass);
    }

    @Override
    public ColumnBuilder<T> column(String name) {
        return this.parentBuilder.column(name);
    }

    @Override
    public ColumnBuilder<T> headerStyle(BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
        return this.headerStyle(null, styleInitializer);
    }

    @Override
    public ColumnBuilder<T> headerStyle(String baseStyleName,
                                        BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
        if (this.headerStyleInitializer != null) {
            throw new SimpleExcelException("Header style for column \"" + this.columnName
                    + "\" has already been defined.");
        }

        this.headerStyleInitializer = this.parentBuilder.parentBuilder.getStyle(baseStyleName, styleInitializer);

        if (this.parentBuilder.headerStyleInitializer != null) {
            this.headerStyleInitializer =
                    this.parentBuilder.headerStyleInitializer.andThen(this.headerStyleInitializer);
        }

        return this;
    }

    @Override
    public ColumnBuilder<T> style(String baseStyleName) {
        return this.style(baseStyleName, null);
    }

    @Override
    public ColumnBuilder<T> style(Consumer<CellStyle> styleInitializer) {
        return this.style(null,
                (CellStyle cellStyle, StyleInitializerContext styleInitializerContext) -> styleInitializer
                        .accept(cellStyle));
    }

    @Override
    public ColumnBuilder<T> style(BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
        return this.style(null, styleInitializer);
    }

    @Override
    public ColumnBuilder<T> style(String baseStyleName,
                                  BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
        if (this.columnStyleInitializer != null) {
            throw new SimpleExcelException("Style for column \"" + this.columnName
                    + "\" has already been defined.");
        }

        this.columnStyleInitializer = this.parentBuilder.parentBuilder.getStyle(baseStyleName, styleInitializer);

        if (this.parentBuilder.cellStyleInitializer != null) {
            this.columnStyleInitializer = this.parentBuilder.cellStyleInitializer.andThen(
                    this.columnStyleInitializer);
        }

        return this;
    }

    @Override
    public ColumnBuilder<T> dataExtractor(Function<T, ?> dataValueExtractor) {
        return dataExtractorWithIndex((dataRow, integer) -> dataValueExtractor.apply(dataRow));
    }

    @Override
    public ColumnBuilder<T> dataExtractorWithIndex(BiFunction<T, Integer, ?> dataValueExtractor) {
        this.dataValueExtractor = dataValueExtractor;
        return this;
    }

    @Override
    public SimpleExcelWriter build() {
        return this.parentBuilder.build();
    }
}
