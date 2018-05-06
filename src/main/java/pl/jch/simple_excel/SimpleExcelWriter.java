package pl.jch.simple_excel;

import lombok.Getter;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;

import static java.util.stream.Collectors.joining;
import static java.util.stream.Collectors.toMap;
import static java.util.stream.Collectors.toSet;
import static org.apache.commons.lang3.Validate.notEmpty;
import static org.apache.commons.lang3.Validate.notNull;

public class SimpleExcelWriter {

    private final SimpleExcelWriterBuilder builder;

    private SimpleExcelWriter(SimpleExcelWriterBuilder builder) {
        this.builder = builder;
    }

    public void write(OutputStream outputStream, DataSet... dataSets) throws IOException {
        final List<DataSet> dataSetList = Arrays.asList(ObjectUtils.defaultIfNull(dataSets, new DataSet[0]));

        InternalWriter.write(builder, dataSetList, outputStream);
    }

    /**
     * This class is intended to serve single workbook write operation.
     */
    private static class InternalWriter {

        private final SimpleExcelWriterBuilder builder;
        private final Map<String, List<?>> dataBySheetName;
        private final OutputStream outputStream;
        private final Workbook workbook;
        private final Map<String, Font> namedFonts = new HashMap<>();
        private final Map<String, CellStyle> namedCellStyles = new HashMap<>();
        private final StyleInitializerContext styleInitializerContext = new StyleInitializerContext() {
            @Override
            public Font definedFont(String fontName) {
                return InternalWriter.this.getNamedFont(fontName);
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

        public InternalWriter(SimpleExcelWriterBuilder builder,
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

        private static void write(SimpleExcelWriterBuilder builder,
                                  List<DataSet> dataSets,
                                  OutputStream outputStream) throws IOException {
            new InternalWriter(builder, dataSets, outputStream).write();
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

        private void createSheet(SheetBuilder sheetBuilder) {
            final Sheet sheet = workbook.createSheet(sheetBuilder.sheetName);

            runSheetCustomizations(sheet, sheetBuilder);

            writeHeader(sheet, sheetBuilder);

            writeData(sheet, sheetBuilder);
        }

        private void writeData(Sheet sheet, SheetBuilder sheetBuilder) {
            final List<?> data = this.dataBySheetName.get(sheet.getSheetName());

            if (data != null) {
                int index = 0;
                for (Object rowData : data) {
                    writeRow(sheet, sheetBuilder, rowData, index);
                    index++;
                }
            }
        }

        private void writeRow(Sheet sheet, SheetBuilder<?> sheetBuilder, Object rowData, int index) {
            final Row row = sheet.createRow(this.nextRowIndex);
            this.nextRowIndex++;

            if (rowData == null) {
                return;
            }

            if (!sheetBuilder.rowClass.isInstance(rowData)) {
                throw new SimpleExcelException("Unexpected data object class for sheet \"" +
                        sheetBuilder.getSheetName() + ". Data index: " + index + ". Expected data class: " +
                        sheetBuilder.rowClass.getName() + ", actual data class: " + rowData.getClass().getName());
            }

            CellStyle[] cellStyles = new CellStyle[sheetBuilder.columnBuilders.size()];

            int i = 0;
            for (ColumnBuilder<?> columnBuilder : sheetBuilder.columnBuilders) {
                if (columnBuilder.columnStyleInitializer != null) {
                    cellStyles[i] = createStyle(columnBuilder.columnStyleInitializer);
                } else if (sheetBuilder.cellStyleInitializer != null) {
                    cellStyles[i] = createStyle(sheetBuilder.cellStyleInitializer);
                }
                i++;
            }

            int cellIndex = 0;
            for (ColumnBuilder columnBuilder : sheetBuilder.columnBuilders) {
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

        private <T> void writeHeader(Sheet sheet, SheetBuilder<?> sheetBuilder) {
            if (sheetBuilder.columnBuilders.isEmpty()) {
                return;
            }

            Row headerRow = sheet.createRow(0);

            final CellStyle[] cellStyles = new CellStyle[sheetBuilder.columnBuilders.size()];
            int j = 0;
            for (ColumnBuilder<?> columnBuilder : sheetBuilder.columnBuilders) {
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
            for (ColumnBuilder<?> columnBuilder : sheetBuilder.columnBuilders) {
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

        private void runSheetCustomizations(Sheet sheet, SheetBuilder<?> sheetBuilder) {
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
                    .map(SheetBuilder::getSheetName)
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

    public void writeToFile(String file, DataSet... dataSets) throws IOException {
        try (OutputStream outputStream = new FileOutputStream(file)) {
            write(outputStream, dataSets);
        }
    }

    public static SimpleExcelWriterBuilder builder() {
        return new SimpleExcelWriterBuilder();
    }

    public static final class SimpleExcelWriterBuilder {

        private WorkbookType workbookType = WorkbookType.HSSF;

        private List<Consumer<Workbook>> workbookCustomizations = new ArrayList<>();
        private List<SheetBuilder> sheetBuilders = new ArrayList<>();
        private Map<String, BiConsumer<CellStyle, StyleInitializerContext>> styleInitializersByName =
                new HashMap<>();
        private Map<String, Consumer<Font>> fontInitializersByName = new HashMap<>();

        private SimpleExcelWriterBuilder() {
        }

        public SimpleExcelWriterBuilder workbookType(WorkbookType workbookType) {
            this.workbookType = notNull(workbookType);
            return this;
        }

        public SimpleExcelWriterBuilder workbookCustomization(Consumer<Workbook> customizer) {
            this.workbookCustomizations.add(customizer);
            return this;
        }

        public SimpleExcelWriterBuilder defineStyle(String name, Consumer<CellStyle> initializer) {
            return defineStyle(name,
                    (CellStyle cellStyle, StyleInitializerContext styleInitializerContext) -> initializer
                            .accept(cellStyle));
        }

        public SimpleExcelWriterBuilder defineStyle(String name,
                                                    BiConsumer<CellStyle, StyleInitializerContext> initializer) {
            if (this.styleInitializersByName.containsKey(name)) {
                throw new SimpleExcelException("Style name already defined: " + name);
            }

            this.styleInitializersByName.put(name, initializer);
            return this;
        }

        public SimpleExcelWriterBuilder defineFont(String name, Consumer<Font> initializer) {
            return this.defineFont(name, null, initializer);
        }

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

        public SheetBuilder<Object> sheet(String name) {
            return sheet(name, Object.class);
        }

        public <T> SheetBuilder<T> sheet(String name, Class<T> rowClass) {
            notEmpty(name, "Name cannot be empty");

            if (isSheetNameDefined(name)) {
                throw new SimpleExcelException("Sheet name already defined: " + name);
            }

            final SheetBuilder<T> result = new SheetBuilder<>(this, name, rowClass);
            this.sheetBuilders.add(result);
            return result;
        }

        private boolean isSheetNameDefined(String name) {
            return this.sheetBuilders.stream()
                    .anyMatch(sheetBuilder -> sheetBuilder.sheetName.equals(name));
        }

        public SimpleExcelWriter build() {
            return new SimpleExcelWriter(this);
        }

        private BiConsumer<CellStyle, StyleInitializerContext> getStyle(String baseStyleName,
                                                                        BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
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

    /**
     * @param <T> row class.
     */
    public static final class SheetBuilder<T> {

        private final SimpleExcelWriterBuilder parentBuilder;
        private final Class<T> rowClass;

        @Getter
        private final String sheetName;
        private final List<Consumer<Sheet>> customizations = new ArrayList<>();
        private final List<ColumnBuilder> columnBuilders = new ArrayList<>();
        private BiConsumer<CellStyle, StyleInitializerContext> headerStyleInitializer;
        private BiConsumer<CellStyle, StyleInitializerContext> cellStyleInitializer;

        public SheetBuilder(SimpleExcelWriterBuilder parentBuilder, String sheetName, Class<T> rowClass) {
            this.parentBuilder = notNull(parentBuilder);
            this.sheetName = notEmpty(sheetName);
            this.rowClass = notNull(rowClass);
        }

        public SheetBuilder<Object> sheet(String name) {
            return sheet(name, Object.class);
        }

        public <S> SheetBuilder<S> sheet(String name, Class<S> rowClass) {
            return this.parentBuilder.sheet(name, rowClass);
        }

        public SheetBuilder<T> sheetCustomization(Consumer<Sheet> customizer) {
            this.customizations.add(customizer);
            return this;
        }

        public ColumnBuilder<T> column(String name) {
            final boolean isNameAlreadyDefined = this.columnBuilders.stream()
                    .map(ColumnBuilder::getColumnName)
                    .anyMatch(s -> s.equals(name));

            if (isNameAlreadyDefined) {
                throw new SimpleExcelException("Column name already defined: " + name);
            }

            final ColumnBuilder<T> result = new ColumnBuilder<>(this, name);
            this.columnBuilders.add(result);
            return result;
        }

        public SheetBuilder<T> headerStyle(String styleName) {
            return this.headerStyle(styleName, null);
        }

        public SheetBuilder<T> headerStyle(Consumer<CellStyle> styleInitializer) {
            return headerStyle(null, (cellStyle, styleInitializerContext) -> styleInitializer.accept(cellStyle));
        }

        public SheetBuilder<T> headerStyle(String baseStyleName,
                                           BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
            if (this.headerStyleInitializer != null) {
                throw new SimpleExcelException("Header style already defined for sheet " + this.sheetName);
            }

            this.headerStyleInitializer = this.parentBuilder.getStyle(baseStyleName, styleInitializer);

            return this;
        }

        public SheetBuilder<T> style(String baseStyleName) {
            return this.style(baseStyleName, null);
        }

        public SheetBuilder<T> style(String baseStyleName,
                                     BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
            if (this.cellStyleInitializer != null) {
                throw new SimpleExcelException("Header style already defined for sheet " + this.sheetName);
            }

            this.cellStyleInitializer = this.parentBuilder.getStyle(baseStyleName, styleInitializer);

            return this;
        }

        public SimpleExcelWriter build() {
            return this.parentBuilder.build();
        }
    }

    /**
     * @param <T> Row class.
     */
    public static class ColumnBuilder<T> {

        private final SheetBuilder<T> parentBuilder;
        @Getter
        private final String columnName;
        private BiConsumer<CellStyle, StyleInitializerContext> headerStyleInitializer;
        private BiConsumer<CellStyle, StyleInitializerContext> columnStyleInitializer;
        private BiFunction<T, Integer, ?> dataValueExtractor;


        private ColumnBuilder(SheetBuilder<T> parentBuilder, String columnName) {
            this.parentBuilder = parentBuilder;
            this.columnName = columnName;
        }

        public ColumnBuilder<T> column(String name) {
            return this.parentBuilder.column(name);
        }

        public SheetBuilder<Object> sheet(String name) {
            return sheet(name, Object.class);
        }

        public <S> SheetBuilder<S> sheet(String name, Class<S> rowClass) {
            return this.parentBuilder.sheet(name, rowClass);
        }

        public ColumnBuilder<T> headerStyle(BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
            return this.headerStyle(null, styleInitializer);
        }

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

        public ColumnBuilder<T> style(String baseStyleName) {
            return this.style(baseStyleName, null);
        }

        public ColumnBuilder<T> style(Consumer<CellStyle> styleInitializer) {
            return this.style(null,
                    (CellStyle cellStyle, StyleInitializerContext styleInitializerContext) -> styleInitializer
                            .accept(cellStyle));
        }

        public ColumnBuilder<T> style(BiConsumer<CellStyle, StyleInitializerContext> styleInitializer) {
            return this.style(null, styleInitializer);
        }

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

        public ColumnBuilder<T> dataExtractor(Function<T, ?> dataValueExtractor) {
            return dataExtractorWithIndex((dataRow, integer) -> dataValueExtractor.apply(dataRow));
        }

        public ColumnBuilder<T> dataExtractorWithIndex(BiFunction<T, Integer, ?> dataValueExtractor) {
            this.dataValueExtractor = dataValueExtractor;
            return this;
        }

        public SimpleExcelWriter build() {
            return this.parentBuilder.build();
        }
    }

}
