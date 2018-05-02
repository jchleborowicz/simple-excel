package pl.jch.simple_excel;

import org.apache.poi.ss.usermodel.Font;

import java.util.function.Consumer;

public interface StyleInitializerContext {

    Font createFont(Consumer<Font> initializer);

    Font definedFont(String fontName);

    short createDataFormat(String format);
}
