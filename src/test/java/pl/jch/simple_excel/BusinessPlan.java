package pl.jch.simple_excel;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class BusinessPlan {

    public static final String CELL_STYLE_HEADER = "header";
    public static final String CELL_STYLE_HEADER_DATE = "header_date";
    public static final String CELL_STYLE_B = "cell_b";
    public static final String CELL_STYLE_B_CENTERED = "cell_b_centered";
    public static final String CELL_STYLE_B_DATE = "cell_b_date";
    public static final String CELL_STYLE_G = "cell_g";
    public static final String CELL_STYLE_BB = "cell_bb";
    public static final String CELL_STYLE_BG = "cell_bg";
    public static final String CELL_STYLE_H = "cell_h";
    public static final String CELL_STYLE_NORMAL = "cell_normal";
    public static final String CELL_STYLE_NORMAL_CENTERED = "cell_normal_centered";
    public static final String CELL_STYLE_NORMAL_DATE = "cell_normal_date";
    public static final String CELL_STYLE_INDENTED = "cell_indented";
    public static final String CELL_STYLE_BLUE = "cell_blue";
    private static SimpleDateFormat fmt = new SimpleDateFormat("dd-MMM");

    private static final String[] TITLES = {"ID", "Project Name", "Owner", "Days", "Start", "End"};

    //sample data to fill the sheet.
    private static final String[][] DATA = {
            {"1.0", "Marketing Research Tactical Plan", "J. Dow", "70", "9-Jul", null,
                    "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"},
            null,
            {"1.1", "Scope Definition Phase", "J. Dow", "10", "9-Jul", null,
                    "x", "x", null, null, null, null, null, null, null, null, null},
            {"1.1.1", "Define research objectives", "J. Dow", "3", "9-Jul", null,
                    "x", null, null, null, null, null, null, null, null, null, null},
            {"1.1.2", "Define research requirements", "S. Jones", "7", "10-Jul", null,
                    "x", "x", null, null, null, null, null, null, null, null, null},
            {"1.1.3", "Determine in-house resource or hire vendor", "J. Dow", "2", "15-Jul", null,
                    "x", "x", null, null, null, null, null, null, null, null, null},
            null,
            {"1.2", "Vendor Selection Phase", "J. Dow", "19", "19-Jul", null,
                    null, "x", "x", "x", "x", null, null, null, null, null, null},
            {"1.2.1", "Define vendor selection criteria", "J. Dow", "3", "19-Jul", null,
                    null, "x", null, null, null, null, null, null, null, null, null},
            {"1.2.2", "Develop vendor selection questionnaire", "S. Jones, T. Wates", "2", "22-Jul", null,
                    null, "x", "x", null, null, null, null, null, null, null, null},
            {"1.2.3", "Develop Statement of Work", "S. Jones", "4", "26-Jul", null,
                    null, null, "x", "x", null, null, null, null, null, null, null},
            {"1.2.4", "Evaluate proposal", "J. Dow, S. Jones", "4", "2-Aug", null,
                    null, null, null, "x", "x", null, null, null, null, null, null},
            {"1.2.5", "Select vendor", "J. Dow", "1", "6-Aug", null,
                    null, null, null, null, "x", null, null, null, null, null, null},
            null,
            {"1.3", "Research Phase", "G. Lee", "47", "9-Aug", null,
                    null, null, null, null, "x", "x", "x", "x", "x", "x", "x"},
            {"1.3.1", "Develop market research information needs questionnaire", "G. Lee", "2", "9-Aug", null,
                    null, null, null, null, "x", null, null, null, null, null, null},
            {"1.3.2", "Interview marketing group for market research needs", "G. Lee", "2", "11-Aug", null,
                    null, null, null, null, "x", "x", null, null, null, null, null},
            {"1.3.3", "Document information needs", "G. Lee, S. Jones", "1", "13-Aug", null,
                    null, null, null, null, null, "x", null, null, null, null, null},
    };

    public static void main(String[] args) throws Exception {
        SimpleExcelWriter.builder()
                .defineStyle("border", (CellStyle style) -> {
                    short blackColorIndex = IndexedColors.BLACK.getIndex();

                    style.setBorderRight(BorderStyle.THIN);
                    style.setRightBorderColor(blackColorIndex);

                    style.setBorderBottom(BorderStyle.THIN);
                    style.setBottomBorderColor(blackColorIndex);

                    style.setBorderLeft(BorderStyle.THIN);
                    style.setLeftBorderColor(blackColorIndex);

                    style.setBorderTop(BorderStyle.THIN);
                    style.setTopBorderColor(blackColorIndex);
                })
                .defineStyle("dateFormat", (CellStyle style, StyleInitializerContext context) ->
                        style.setDataFormat(context.createDataFormat("d-mmm")))
                .defineFont("bold", (Font font) -> font.setBold(true))
                .defineFont("boldItalic", "bold", (Font font) -> font.setItalic(true))

                .sheet("Business Plan", Object[].class)
                .sheetCustomization(BusinessPlan::customizeSheet)
                .headerStyle("border", (CellStyle style, StyleInitializerContext context) -> {
                    style.setFont(context.definedFont("bold"));
                    style.setAlignment(HorizontalAlignment.CENTER);
                    style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                })
                .style("border")
                .column("ID")
                .headerStyle((CellStyle cellStyle, StyleInitializerContext context) ->
                        cellStyle.setFont(context.definedFont("boldItalic")))
                .dataExtractor(objects -> objects[0])
                .column("Project Name")
                .style((CellStyle cellStyle, StyleInitializerContext context) ->
                        cellStyle.setFont(context.definedFont("bold")))
                .dataExtractor(objects -> objects[1])
                .column("Owner")
                .dataExtractor(objects -> objects[2])
                .column("Days")
                .dataExtractor(objects -> objects[3])
                .column("Start")
                .dataExtractor(objects -> objects[4])
                .column("End")
                .dataExtractor(objects -> objects[5])

                .sheet("Old Business Plan")
                .sheetCustomization(sheet -> {
                    try {
                        process(sheet);
                    } catch (ParseException e) {
                        throw new RuntimeException(e);
                    }
                })
                .sheetCustomization(BusinessPlan::customizeSheet)
                .build()

                .writeToFile("businessplan.xlsx", DataSet.of("Business Plan", Arrays.asList(DATA)));
    }

    private static Map<String, CellStyle> createStyles(Workbook workbook) {
        Map<String, CellStyle> styles = new HashMap<>();

        createHeaderStyle(workbook, styles);

        createHeaderDateStyle(workbook, styles);

        Font font1 = workbook.createFont();
        font1.setBold(true);

        CellStyle bStyle = createBorderedStyle(workbook);
        bStyle.setAlignment(HorizontalAlignment.LEFT);
        bStyle.setFont(font1);
        styles.put(CELL_STYLE_B, bStyle);

        CellStyle bCenteredStyle = createBorderedStyle(workbook);
        bCenteredStyle.setAlignment(HorizontalAlignment.CENTER);
        bCenteredStyle.setFont(font1);
        styles.put(CELL_STYLE_B_CENTERED, bCenteredStyle);

        CellStyle bDateStyle = createBorderedStyle(workbook);
        bDateStyle.setAlignment(HorizontalAlignment.RIGHT);
        bDateStyle.setFont(font1);
        bDateStyle.setDataFormat(workbook.createDataFormat().getFormat("d-mmm"));
        styles.put(CELL_STYLE_B_DATE, bCenteredStyle);

        CellStyle gStyle = createBorderedStyle(workbook);
        gStyle.setAlignment(HorizontalAlignment.RIGHT);
        gStyle.setFont(font1);
        gStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        gStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        gStyle.setDataFormat(workbook.createDataFormat().getFormat("d-mmm"));
        styles.put(CELL_STYLE_G, gStyle);

        Font font2 = workbook.createFont();
        font2.setColor(IndexedColors.BLUE.getIndex());
        font2.setBold(true);

        CellStyle bbStyle = createBorderedStyle(workbook);
        bbStyle.setAlignment(HorizontalAlignment.LEFT);
        bbStyle.setFont(font2);
        styles.put(CELL_STYLE_BB, bbStyle);

        CellStyle bgStyle = createBorderedStyle(workbook);
        bgStyle.setAlignment(HorizontalAlignment.RIGHT);
        bgStyle.setFont(font1);
        bgStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        bgStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        bgStyle.setDataFormat(workbook.createDataFormat().getFormat("d-mmm"));
        styles.put(CELL_STYLE_BG, bgStyle);

        Font font3 = workbook.createFont();
        font3.setFontHeightInPoints((short) 14);
        font3.setColor(IndexedColors.DARK_BLUE.getIndex());
        font3.setBold(true);

        CellStyle hStyle = createBorderedStyle(workbook);
        hStyle.setAlignment(HorizontalAlignment.LEFT);
        hStyle.setFont(font3);
        hStyle.setWrapText(true);
        styles.put(CELL_STYLE_H, hStyle);

        CellStyle normalStyle = createBorderedStyle(workbook);
        normalStyle.setAlignment(HorizontalAlignment.LEFT);
        normalStyle.setWrapText(true);
        styles.put(CELL_STYLE_NORMAL, normalStyle);

        CellStyle normalCenteredStyle = createBorderedStyle(workbook);
        normalCenteredStyle.setAlignment(HorizontalAlignment.CENTER);
        normalCenteredStyle.setWrapText(true);
        styles.put(CELL_STYLE_NORMAL_CENTERED, normalCenteredStyle);

        CellStyle normalDateStyle = createBorderedStyle(workbook);
        normalDateStyle.setAlignment(HorizontalAlignment.RIGHT);
        normalDateStyle.setWrapText(true);
        normalDateStyle.setDataFormat(workbook.createDataFormat().getFormat("d-mmm"));
        styles.put(CELL_STYLE_NORMAL_DATE, normalDateStyle);

        CellStyle indentedStyle = createBorderedStyle(workbook);
        indentedStyle.setAlignment(HorizontalAlignment.LEFT);
        indentedStyle.setIndention((short) 1);
        indentedStyle.setWrapText(true);
        styles.put(CELL_STYLE_INDENTED, indentedStyle);

        CellStyle blueStyle = createBorderedStyle(workbook);
        blueStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        blueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put(CELL_STYLE_BLUE, blueStyle);

        return styles;
    }

    private static void createHeaderDateStyle(Workbook workbook, Map<String, CellStyle> styles) {
        CellStyle headerDateStyle = createBorderedStyle(workbook);
        headerDateStyle.setAlignment(HorizontalAlignment.CENTER);
        headerDateStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        headerDateStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerDateStyle.setFont(createHeaderFont(workbook));
        headerDateStyle.setDataFormat(workbook.createDataFormat().getFormat("d-mmm"));
        styles.put(CELL_STYLE_HEADER_DATE, headerDateStyle);
    }

    private static void createHeaderStyle(Workbook workbook, Map<String, CellStyle> styles) {
        CellStyle headerStyle = createBorderedStyle(workbook);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setFont(createHeaderFont(workbook));
        styles.put(CELL_STYLE_HEADER, headerStyle);
    }

    private static Font createHeaderFont(Workbook workbook) {
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        return headerFont;
    }

    private static void process(Sheet sheet) throws ParseException {
        final Map<String, CellStyle> styles = createStyles(sheet.getWorkbook());

        //the header row: centered text in 48pt font
        Row headerRow = sheet.createRow(0);
        headerRow.setHeightInPoints(12.75f);
        for (int i = 0; i < TITLES.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(TITLES[i]);
            cell.setCellStyle(styles.get(CELL_STYLE_HEADER));
        }
        //columns for 11 weeks starting from 9-Jul
        Calendar calendar = Calendar.getInstance();
        int year = calendar.get(Calendar.YEAR);

        calendar.setTime(fmt.parse("9-Jul"));
        calendar.set(Calendar.YEAR, year);
        for (int i = 0; i < 11; i++) {
            Cell cell = headerRow.createCell(TITLES.length + i);
            cell.setCellValue(calendar);
            cell.setCellStyle(styles.get(CELL_STYLE_HEADER_DATE));
            calendar.roll(Calendar.WEEK_OF_YEAR, true);
        }
        //freeze the first row
        sheet.createFreezePane(0, 1);

        Row row;
        Cell cell;
        int rownum = 1;
        for (int i = 0; i < DATA.length; i++, rownum++) {
            row = sheet.createRow(rownum);
            if (DATA[i] == null) {
                continue;
            }

            for (int columnIndex = 0; columnIndex < DATA[i].length; columnIndex++) {
                cell = row.createCell(columnIndex);
                String styleName;
                boolean isHeader = i == 0 || DATA[i - 1] == null;
                switch (columnIndex) {
                    case 0:
                        if (isHeader) {
                            styleName = CELL_STYLE_B;
                            cell.setCellValue(Double.parseDouble(DATA[i][columnIndex]));
                        } else {
                            styleName = CELL_STYLE_NORMAL;
                            cell.setCellValue(DATA[i][columnIndex]);
                        }
                        break;
                    case 1:
                        if (isHeader) {
                            styleName = i == 0 ? "cell_h" : CELL_STYLE_BB;
                        } else {
                            styleName = "cell_indented";
                        }
                        cell.setCellValue(DATA[i][columnIndex]);
                        break;
                    case 2:
                        styleName = isHeader ? CELL_STYLE_B : CELL_STYLE_NORMAL;
                        cell.setCellValue(DATA[i][columnIndex]);
                        break;
                    case 3:
                        styleName = isHeader ? CELL_STYLE_B_CENTERED : CELL_STYLE_NORMAL_CENTERED;
                        cell.setCellValue(Integer.parseInt(DATA[i][columnIndex]));
                        break;
                    case 4: {
                        calendar.setTime(fmt.parse(DATA[i][columnIndex]));
                        calendar.set(Calendar.YEAR, year);
                        cell.setCellValue(calendar);
                        styleName = isHeader ? CELL_STYLE_B_DATE : "cell_normal_date";
                        break;
                    }
                    case 5: {
                        int r = rownum + 1;
                        String fmla = "IF(AND(D" + r + ",E" + r + "),E" + r + "+D" + r + ",\"\")";
                        cell.setCellFormula(fmla);
                        styleName = isHeader ? CELL_STYLE_BG : "cell_g";
                        break;
                    }
                    default:
                        styleName = DATA[i][columnIndex] != null ? "cell_blue" : CELL_STYLE_NORMAL;
                }

                cell.setCellStyle(styles.get(styleName));
            }
        }

        //group rows for each phase, row numbers are 0-based
        sheet.groupRow(4, 6);
        sheet.groupRow(9, 13);
        sheet.groupRow(16, 18);

        //set column widths, the width is measured in units of 1/256th of a character width
        sheet.setColumnWidth(0, 256 * 6);
        sheet.setColumnWidth(1, 256 * 33);
        sheet.setColumnWidth(2, 256 * 20);
        sheet.setZoom(75); //75% scale
    }

    private static void customizeSheet(Sheet sheet) {
        //turn off gridlines
        sheet.setDisplayGridlines(false);
        sheet.setPrintGridlines(false);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);

        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);

        //the following three statements are required only for HSSF
        sheet.setAutobreaks(true);
        printSetup.setFitHeight((short) 1);
        printSetup.setFitWidth((short) 1);
    }

    static CellStyle createBorderedStyle(Workbook workbook) {
        final CellStyle style = workbook.createCellStyle();

        short blackColorIndex = IndexedColors.BLACK.getIndex();

        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(blackColorIndex);

        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(blackColorIndex);

        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(blackColorIndex);

        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(blackColorIndex);

        return style;
    }
}
