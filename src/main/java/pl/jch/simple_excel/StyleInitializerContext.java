package pl.jch.simple_excel;

import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.Font;

public interface StyleInitializerContext {

    Font createFont(Consumer<Font> initializer);

    Font definedFont(String fontName);

    short createDataFormat(String format);
}
