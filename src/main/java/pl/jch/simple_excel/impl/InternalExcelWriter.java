package pl.jch.simple_excel.impl;

import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import pl.jch.simple_excel.DataSet;
import pl.jch.simple_excel.SimpleExcelException;
import pl.jch.simple_excel.StyleInitializerContext;
import pl.jch.simple_excel.WorkbookType;

import static java.util.stream.Collectors.joining;
import static java.util.stream.Collectors.toMap;
import static java.util.stream.Collectors.toSet;
import static org.apache.commons.lang3.Validate.notNull;


/**
 * This class is intended to serve single workbook write operation.
 */
class InternalExcelWriter {

    private final SimpleExcelWriterBuilderImpl builder;
    private final Map<String, List<?>> dataBySheetName;
    private final OutputStream outputStream;
    private final Workbook workbook;
    private final Map<String, Font> namedFonts = new HashMap<>();
    private final Map<String, CellStyle> namedCellStyles = new HashMap<>();
    private final StyleInitializerContext styleInitializerContext = new StyleInitializerContext() {
        @Override
        public Font definedFont(String fontName) {
            return InternalExcelWriter.this.getNamedFont(fontName);
        }

        @Override
        public Font createFont(Consumer<Font> initializer) {
            final Font result = workbook.createFont();
            if (initializer != null) {
                initializer.accept(result);
            }
            return result;
        }

        @Override
        public short createDataFormat(String format) {
            return workbook.createDataFormat().getFormat(format);
        }
    };
    private int nextRowIndex = 1;

    public InternalExcelWriter(SimpleExcelWriterBuilderImpl builder,
                               List<DataSet> dataSets,
                               OutputStream outputStream) {
        this.builder = notNull(builder);
        this.dataBySheetName = dataSets.stream()
                .collect(toMap(DataSet::getSheetName, DataSet::getData));
        this.outputStream = notNull(outputStream);

        this.workbook = createWorkbook(builder.workbookType);
    }

    private static Workbook createWorkbook(WorkbookType workbookType) {
        switch (workbookType) {
            case HSSF:
                return new HSSFWorkbook();
            case XSSF:
                return new XSSFWorkbook();
            default:
                throw new SimpleExcelException("Unsupported workbook type: " + workbookType);
        }
    }

    public static void write(SimpleExcelWriterBuilderImpl builder,
                             List<DataSet> dataSets,
                             OutputStream outputStream) throws IOException {
        new InternalExcelWriter(builder, dataSets, outputStream).write();
    }

    private void write() throws IOException {
        verifyDataSheetsExist();

        applyWorkbookCustomizations();

        createSheets();

        writeWorkbookToOutput();
    }

    private void createSheets() {
        this.builder.sheetBuilders.forEach(this::createSheet);
    }

    private void writeWorkbookToOutput() throws IOException {
        workbook.write(outputStream);
        workbook.close();
    }

    private void createSheet(SheetBuilderImpl sheetBuilder) {
        final Sheet sheet = workbook.createSheet(sheetBuilder.sheetName);

        runSheetCustomizations(sheet, sheetBuilder);

        writeHeader(sheet, sheetBuilder);

        writeData(sheet, sheetBuilder);
    }

    private void writeData(Sheet sheet, SheetBuilderImpl sheetBuilder) {
        final List<?> data = this.dataBySheetName.get(sheet.getSheetName());

        if (data != null) {
            int index = 0;
            for (Object rowData : data) {
                writeRow(sheet, sheetBuilder, rowData, index);
                index++;
            }
        }
    }

    private void writeRow(Sheet sheet, SheetBuilderImpl<?> sheetBuilder, Object rowData, int index) {
        final Row row = sheet.createRow(this.nextRowIndex);
        this.nextRowIndex++;

        if (rowData == null) {
            return;
        }

        if (!sheetBuilder.rowClass.isInstance(rowData)) {
            throw new SimpleExcelException("Unexpected data object class for sheet \""
                    + sheetBuilder.getSheetName() + ". Data index: " + index + ". Expected data class: "
                    + sheetBuilder.rowClass.getName() + ", actual data class: " + rowData.getClass().getName());
        }

        CellStyle[] cellStyles = new CellStyle[sheetBuilder.columnBuilders.size()];

        int i = 0;
        for (ColumnBuilderImpl<?> columnBuilder : sheetBuilder.columnBuilders) {
            if (columnBuilder.columnStyleInitializer != null) {
                cellStyles[i] = createStyle(columnBuilder.columnStyleInitializer);
            } else if (sheetBuilder.cellStyleInitializer != null) {
                cellStyles[i] = createStyle(sheetBuilder.cellStyleInitializer);
            }
            i++;
        }

        int cellIndex = 0;
        for (ColumnBuilderImpl columnBuilder : sheetBuilder.columnBuilders) {
            final Cell cell = row.createCell(cellIndex);

            @SuppressWarnings("unchecked") final Object value =
                    columnBuilder.dataValueExtractor.apply(rowData, index);

            //todo asap set correct value
            if (value != null) {
                cell.setCellValue(value.toString());
            }

            if (cellStyles[cellIndex] != null) {
                cell.setCellStyle(cellStyles[cellIndex]);
            }

            cellIndex++;
        }
    }

    private <T> void writeHeader(Sheet sheet, SheetBuilderImpl<?> sheetBuilder) {
        if (sheetBuilder.columnBuilders.isEmpty()) {
            return;
        }

        Row headerRow = sheet.createRow(0);

        final CellStyle[] cellStyles = new CellStyle[sheetBuilder.columnBuilders.size()];
        int j = 0;
        for (ColumnBuilderImpl<?> columnBuilder : sheetBuilder.columnBuilders) {
            if (columnBuilder.headerStyleInitializer != null) {
                cellStyles[j] = createStyle(columnBuilder.headerStyleInitializer);
            } else if (sheetBuilder.headerStyleInitializer != null) {
                cellStyles[j] = createStyle(sheetBuilder.headerStyleInitializer);
            } else if (columnBuilder.columnStyleInitializer != null) {
                cellStyles[j] = createStyle(columnBuilder.columnStyleInitializer);
            }
            j++;
        }

        int i = 0;
        for (ColumnBuilderImpl<?> columnBuilder : sheetBuilder.columnBuilders) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columnBuilder.columnName);

            if (cellStyles[i] != null) {
                cell.setCellStyle(cellStyles[i]);
            }
            i++;
        }
    }

    private CellStyle createStyle(BiConsumer<CellStyle, StyleInitializerContext> initializer) {
        final CellStyle result = this.workbook.createCellStyle();
        initializer.accept(result, this.styleInitializerContext);
        return result;
    }

    private void runSheetCustomizations(Sheet sheet, SheetBuilderImpl<?> sheetBuilder) {
        try {
            sheetBuilder.customizations.forEach(customization -> customization.accept(sheet));
        } catch (RuntimeException e) {
            throw new SimpleExcelException("Error when customizing sheet " + sheetBuilder.sheetName);
        }
    }

    private Font getNamedFont(String fontName) {
        if (!this.namedFonts.containsKey(fontName)) {
            final Font createdFont = createNamedFont(fontName);

            this.namedFonts.put(fontName, createdFont);
        }
        return this.namedFonts.get(fontName);
    }

    private CellStyle getNamedStyle(String styleName) {
        if (!this.namedCellStyles.containsKey(styleName)) {
            final CellStyle createdStyle = createNamedStyle(styleName);
            this.namedCellStyles.put(styleName, createdStyle);
        }
        return this.namedCellStyles.get(styleName);
    }

    private Font createNamedFont(String fontName) {
        final Consumer<Font> initializer = this.builder.fontInitializersByName.get(fontName);

        if (initializer == null) {
            throw new SimpleExcelException("Font name has not been defined: " + fontName);
        }

        final Font result = this.workbook.createFont();
        initializer.accept(result);

        return result;
    }

    private CellStyle createNamedStyle(String styleName) {
        final BiConsumer<CellStyle, StyleInitializerContext> initializer =
                this.builder.styleInitializersByName.get(styleName);

        if (initializer == null) {
            throw new SimpleExcelException("Style has not been defined: " + styleName);
        }

        return createStyle(initializer);
    }

    private void applyWorkbookCustomizations() {
        this.builder.workbookCustomizations.forEach(
                workbookCustomization -> workbookCustomization.accept(this.workbook));
    }

    private void verifyDataSheetsExist() {
        final Set<String> definedSheetNames = this.builder.sheetBuilders.stream()
                .map(SheetBuilderImpl::getSheetName)
                .collect(toSet());

        final String nonExistingSheetNames = this.dataBySheetName.keySet()
                .stream()
                .filter(sheetName -> !definedSheetNames.contains(sheetName))
                .collect(joining(", "));

        if (!nonExistingSheetNames.isEmpty()) {
            throw new SimpleExcelException("No sheets with names: " + nonExistingSheetNames);
        }
    }

}
