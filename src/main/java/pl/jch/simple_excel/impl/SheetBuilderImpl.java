package pl.jch.simple_excel.impl;

import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import pl.jch.simple_excel.ColumnBuilder;
import pl.jch.simple_excel.SheetBuilder;
import pl.jch.simple_excel.SimpleExcelException;
import pl.jch.simple_excel.SimpleExcelWriter;
import pl.jch.simple_excel.StyleInitializerContext;

import static org.apache.commons.lang3.Validate.notEmpty;
import static org.apache.commons.lang3.Validate.notNull;

/**
 * @param <T> row class.
 */
public final class SheetBuilderImpl<T> implements SheetBuilder<T> {

    final SimpleExcelWriterBuilderImpl parentBuilder;
    final Class<T> rowClass;

    @Getter
    final String sheetName;
    final List<Consumer<Sheet>> customizations = new ArrayList<>();
    final List<ColumnBuilderImpl> columnBuilders = new ArrayList<>();
    BiConsumer<CellStyle, StyleInitializerContext> headerStyleInitializer;
    BiConsumer<CellStyle, StyleInitializerContext> cellStyleInitializer;

    public SheetBuilderImpl(SimpleExcelWriterBuilderImpl parentBuilder, String sheetName, Class<T> rowClass) {
        this.parentBuilder = notNull(parentBuilder);
        this.sheetName = notEmpty(sheetName);
        this.rowClass = notNull(rowClass);
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
    public SheetBuilder<T> sheetCustomization(Consumer<Sheet> customizer) {
        this.customizations.add(customizer);
        return this;
    }

    @Override
    public ColumnBuilder<T> column(String name) {
        final boolean isNameAlreadyDefined = this.columnBuilders.stream()
                .map(ColumnBuilderImpl::getColumnName)
                .anyMatch(s -> s.equals(name));

        if (isNameAlreadyDefined) {
            throw new SimpleExcelException("Column name already defined: " + name);
        }

        final ColumnBuilderImpl<T> result = new ColumnBuilderImpl<>(this, name);
        this.columnBuilders.add(result);
        return result;
    }

    @Override
    public SheetBuilder<T> headerStyle(String styleName) {
        return this.headerStyle(styleName, null);
    }

    @Override
    public SheetBuilder<T> headerStyle(Consumer<CellStyle> styleInitializer) {
        return headerStyle(null, (cellStyle, styleInitializerContext) -> styleInitializer.accept(cellStyle));
    }

    @Override
    public SheetBuilder<T> headerStyle(String baseStyleName,
                                       BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
        if (this.headerStyleInitializer != null) {
            throw new SimpleExcelException("Header style already defined for sheet " + this.sheetName);
        }

        this.headerStyleInitializer = this.parentBuilder.getStyle(baseStyleName, styleInitializer);

        return this;
    }

    @Override
    public SheetBuilder<T> style(String baseStyleName) {
        return this.style(baseStyleName, null);
    }

    @Override
    public SheetBuilder<T> style(String baseStyleName,
                                 BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
        if (this.cellStyleInitializer != null) {
            throw new SimpleExcelException("Header style already defined for sheet " + this.sheetName);
        }

        this.cellStyleInitializer = this.parentBuilder.getStyle(baseStyleName, styleInitializer);

        return this;
    }

    @Override
    public SimpleExcelWriter build() {
        return this.parentBuilder.build();
    }
}
