package pl.jch.simple_excel.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import pl.jch.simple_excel.SheetBuilder;
import pl.jch.simple_excel.SimpleExcelException;
import pl.jch.simple_excel.SimpleExcelWriter;
import pl.jch.simple_excel.SimpleExcelWriterBuilder;
import pl.jch.simple_excel.StyleInitializerContext;
import pl.jch.simple_excel.WorkbookType;

import static org.apache.commons.lang3.Validate.notEmpty;
import static org.apache.commons.lang3.Validate.notNull;

public final class SimpleExcelWriterBuilderImpl implements SimpleExcelWriterBuilder {

    WorkbookType workbookType = WorkbookType.HSSF;
    List<Consumer<Workbook>> workbookCustomizations = new ArrayList<>();
    List<SheetBuilderImpl> sheetBuilders = new ArrayList<>();
    Map<String, BiConsumer<CellStyle, StyleInitializerContext>> styleInitializersByName =
            new HashMap<>();
    Map<String, Consumer<Font>> fontInitializersByName = new HashMap<>();

    public SimpleExcelWriterBuilderImpl() {
    }

    @Override
    public SimpleExcelWriterBuilder workbookType(WorkbookType workbookType) {
        this.workbookType = notNull(workbookType);
        return this;
    }

    @Override
    public SimpleExcelWriterBuilder workbookCustomization(Consumer<Workbook> customizer) {
        this.workbookCustomizations.add(customizer);
        return this;
    }

    @Override
    public SimpleExcelWriterBuilder defineStyle(String name, Consumer<CellStyle> initializer) {
        return defineStyle(name,
                (CellStyle cellStyle, StyleInitializerContext styleInitializerContext) -> initializer
                        .accept(cellStyle));
    }

    @Override
    public SimpleExcelWriterBuilder defineStyle(String name,
                                                BiConsumer<CellStyle, StyleInitializerContext> initializer) {
        if (this.styleInitializersByName.containsKey(name)) {
            throw new SimpleExcelException("Style name already defined: " + name);
        }

        this.styleInitializersByName.put(name, initializer);
        return this;
    }

    @Override
    public SimpleExcelWriterBuilder defineFont(String name, Consumer<Font> initializer) {
        return this.defineFont(name, null, initializer);
    }

    @Override
    public SimpleExcelWriterBuilder defineFont(String name, String baseFontName, Consumer<Font> initializer) {
        if (this.fontInitializersByName.containsKey(name)) {
            throw new SimpleExcelException("Font name already defined: " + name);
        }

        final Consumer<Font> effectiveInitializer;

        if (baseFontName == null) {
            effectiveInitializer = initializer;
        } else {
            final Consumer<Font> baseFontInitializer = this.fontInitializersByName.get(baseFontName);
            if (baseFontInitializer == null) {
                throw new SimpleExcelException("Cannt find base font with name: " + baseFontName);
            }

            effectiveInitializer = baseFontInitializer.andThen(initializer);
        }

        this.fontInitializersByName.put(name, effectiveInitializer);

        return this;
    }

    @Override
    public SheetBuilder<Object> sheet(String name) {
        return sheet(name, Object.class);
    }

    @Override
    public <T> SheetBuilder<T> sheet(String name, Class<T> rowClass) {
        notEmpty(name, "Name cannot be empty");

        if (isSheetNameDefined(name)) {
            throw new SimpleExcelException("Sheet name already defined: " + name);
        }

        final SheetBuilderImpl<T>
                result = new SheetBuilderImpl<>(this, name, rowClass);
        this.sheetBuilders.add(result);
        return result;
    }

    public boolean isSheetNameDefined(String name) {
        return this.sheetBuilders.stream()
                .anyMatch(sheetBuilder -> sheetBuilder.getSheetName().equals(name));
    }

    @Override
    public SimpleExcelWriter build() {
        return new SimpleExcelWriterImpl(this);
    }

    BiConsumer<CellStyle, StyleInitializerContext> getStyle(String baseStyleName,
                                                            BiConsumer<CellStyle,
                                                                    StyleInitializerContext> styleInitializer) {
        if (baseStyleName != null) {
            final BiConsumer<CellStyle, StyleInitializerContext> baseInitializer =
                    this.styleInitializersByName.get(baseStyleName);

            if (baseInitializer == null) {
                throw new SimpleExcelException("Base style not found. Base style name: " + baseStyleName);
            }

            return styleInitializer == null ? baseInitializer : baseInitializer.andThen(styleInitializer);
        } else if (styleInitializer != null) {
            return styleInitializer;
        } else {
            throw new SimpleExcelException("Either base style name or style initializer must be specified");
        }
    }
}
