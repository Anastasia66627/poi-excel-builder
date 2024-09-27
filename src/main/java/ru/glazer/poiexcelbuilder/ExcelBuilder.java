package ru.glazer.poiexcelbuilder;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xddf.usermodel.text.XDDFTextBody;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTrendline;
import org.openxmlformats.schemas.drawingml.x2006.chart.STTrendlineType;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineWidth;
import org.openxmlformats.schemas.drawingml.x2006.main.STPresetLineDashVal;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.NoSuchElementException;

public class ExcelBuilder {

    public XSSFWorkbook createWorkbookEnvironment() {
        return new XSSFWorkbook();
    }

    public XSSFSheet createSheet(XSSFWorkbook workbook, String sheetName) {
        if (workbook == null) {
            throw new NoSuchElementException("Workbook does not exist");
        }
        XSSFSheet sheet;
        if (sheetName == null) {
            sheet = workbook.createSheet();
        } else {
            sheet = workbook.createSheet(sheetName);
        }

        return sheet;
    }

    public void createPrintSetup(XSSFSheet sheet, boolean landscape, String repeatingRowsRange, PaperSize paperSize,
                                 boolean fitToPage, short fitHeight, short fitWidth) {
        if (sheet == null) {
            throw new NoSuchElementException("Sheet does not exist");
        }
        XSSFPrintSetup printSetup = sheet.getPrintSetup();
        // устанавливает альбомную ориентацию при печати
        printSetup.setLandscape(landscape);
        // устанавливает повторяющиеся строки вначале каждой страницы для печати.
        // ИСПОЛЬЗОВАТЬ ТОЛЬКО ПОСЛЕ МЕТОДА sheet.getPrintSetup().setLandscape()!!!
        if (repeatingRowsRange != null) {
            sheet.setRepeatingRows(CellRangeAddress.valueOf(repeatingRowsRange));
        }
        // устанавливает размер страницы
        printSetup.setPaperSize(paperSize);
        // подгоняет ширину столбцов под ширину страницы
        sheet.setFitToPage(fitToPage);
        // устанавливает максимальное количество страниц вмещенных на страницу по высоте
        printSetup.setFitHeight(fitHeight);
        // устанавливает максимальное количество страниц вмещенных на страницу по ширине
        printSetup.setFitWidth(fitWidth);
    }

    public XSSFRow createCellByRow(XSSFSheet sheet, int rowIndex, int cellIndex, Object cellValue, XSSFCellStyle style) {
        if (sheet == null) {
            throw new NoSuchElementException("Sheet does not exist");
        }

        XSSFRow row = sheet.getRow(rowIndex) == null ? sheet.createRow(rowIndex) : sheet.getRow(rowIndex);
        XSSFCell cell = row.createCell(cellIndex);

        //TODO: for 21 java
//        switch (cellValue) {
//            case String s -> cell.setCellValue(s);
//            case Double d -> cell.setCellValue(d);
//            case Boolean b -> cell.setCellValue(b);
//            case Date date -> cell.setCellValue(date);
//            case Calendar c -> cell.setCellValue(c);
//            case LocalDate ld -> cell.setCellValue(ld);
//            case LocalDateTime ldt -> cell.setCellValue(ldt);
//            case RichTextString rts -> cell.setCellValue(rts);
//            default -> cell.setCellValue((int)cellValue);
//        }

        if (cellValue instanceof String) {
            cell.setCellValue((String) cellValue);
        } else if (cellValue instanceof Double) {
            cell.setCellValue((double) cellValue);
        } else if (cellValue instanceof Boolean) {
            cell.setCellValue((boolean) cellValue);
        } else if (cellValue instanceof Date) {
            cell.setCellValue((Date) cellValue);
        } else if (cellValue instanceof Calendar) {
            cell.setCellValue((Calendar) cellValue);
        } else if (cellValue instanceof LocalDate) {
            cell.setCellValue((LocalDate) cellValue);
        } else if (cellValue instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) cellValue);
        } else if (cellValue instanceof RichTextString) {
            cell.setCellValue((RichTextString) cellValue);
        } else {
            cell.setCellValue((int)cellValue);
        }

        cell.setCellStyle(style);

        return row;
    }

    public void createRowCellRange(XSSFSheet sheet, int rowIndex, int fromCell, int toCell, XSSFCellStyle style) {
        if (sheet == null) {
            throw new NoSuchElementException("Sheet does not exist");
        }

        XSSFRow row = sheet.getRow(rowIndex) == null ? sheet.createRow(rowIndex) : sheet.getRow(rowIndex);
        for (int i = fromCell; i <= toCell; i++) {
            row.createCell(i).setCellStyle(style);
        }
    }

    public void createColumnCellRange(XSSFSheet sheet, int columnIndex, int fromRow, int toRow, XSSFCellStyle style) {
        if (sheet == null) {
            throw new NoSuchElementException("Sheet does not exist");
        }
        // создает строки если не существуют
        for (int i = fromRow; i <= toRow; i++) {
            XSSFRow row = sheet.getRow(fromRow) == null ? sheet.createRow(fromRow) : sheet.getRow(fromRow);
            row.createCell(columnIndex).setCellStyle(style);
        }
    }

    public XSSFCellStyle createNewStyle(XSSFWorkbook workbook, String fontName, double height, boolean bold,
                                        HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment,
                                        boolean setBorders, BorderStyle borderStyle, boolean wrapText) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(horizontalAlignment);
        style.setVerticalAlignment(verticalAlignment);

        if (setBorders) {
            style.setBorderBottom(borderStyle);
            style.setBorderTop(borderStyle);
            style.setBorderRight(borderStyle);
            style.setBorderLeft(borderStyle);
        }
        style.setWrapText(wrapText);

        XSSFFont font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeight(height);
        font.setBold(bold);

        style.setFont(font);

        return style;
    }

    public XSSFChart createChart(XSSFSheet sheet, int fCellX, int fCellY, int sCellX, int sCellY,
                                 int fCellColumn, int fCellRow, int sCellColumn, int sCellRow,
                                 String title, String chartHeaderFont, int titleSize) {
        if (sheet == null) {
            throw new NoSuchElementException("Sheet does not exist");
        }
        // создание чертежа/рисунка/эскиза листа
        XSSFDrawing drawingPatriarch = sheet.createDrawingPatriarch();
        // создание привязки для графика. Указываются координаты x,y в первой и во второй ячейках, а также их позиции на листе (колонки и строки)
        XSSFClientAnchor anchor = drawingPatriarch.createAnchor(fCellX, fCellY, sCellX, sCellY, fCellColumn, fCellRow, sCellColumn, sCellRow);
        // создание графика с помощью Anchor
        XSSFChart chart = drawingPatriarch.createChart(anchor);
        chart.setTitleText(title);

//        // установка шрифта и его размера (1400) для заголовка графика
        chart.getCTChart().getTitle().getTx().getRich().getPArray(0).getRArray(0).getRPr().setSz(titleSize);
        chart.getCTChart().getTitle().getTx().getRich().getPArray(0).getRArray(0).getRPr().addNewLatin().setTypeface(chartHeaderFont);

        return chart;
    }

    public void createChartLegend(XSSFChart chart, LegendPosition legendPosition, double fontSize) {
        if (chart == null) {
            throw new NoSuchElementException("Chart does not exist");
        }
        // создание легенды графика и указание ее позиции
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(legendPosition);

        // создание тела легенды и установка размера шрифта
        XDDFTextBody legendTextBody = new XDDFTextBody(legend);
        legendTextBody.getXmlObject().addNewBodyPr();
        legendTextBody.addNewParagraph().addDefaultRunProperties().setFontSize(fontSize);
        legend.setTextBody(legendTextBody);
    }

    public XDDFChartData createChartData(XSSFChart chart, ChartTypes chartType, AxisPosition xPosition, AxisTickMark xMark,
                                         double xFontSize, AxisPosition yPosition, AxisTickMark yMark, double yFontSize,
                                         AxisCrosses axisCrosses, AxisCrossBetween axisCrossBetween) {
        if (chart == null) {
            throw new NoSuchElementException("Chart does not exist");
        }
        // создание оси x (ось категорий) и указание ее позиции на графике (BOTTOM)
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(xPosition);
        // отметки на оси не отображаются (NONE)
        bottomAxis.setMajorTickMark(xMark);
        // 9.0
        bottomAxis.getOrAddTextProperties().setFontSize(xFontSize);

        // создание оси y (ось значений) и указание ее позиции на графике (LEFT)
        XDDFValueAxis leftAxis = chart.createValueAxis(yPosition);
        // положение пересечения осей (AUTO_ZERO)
        leftAxis.setCrosses(axisCrosses);
        // устанавливает положение оси между делениями (BETWEEN)
        leftAxis.setCrossBetween(axisCrossBetween);
        // (NONE)
        leftAxis.setMajorTickMark(yMark);
        leftAxis.getOrAddTextProperties().setFontSize(yFontSize);

        return chart.createData(chartType, bottomAxis, leftAxis);
    }

    public XDDFChartData.Series createDataSeriesChartFromArray(XDDFChartData data, String[] categories, Integer[] values,
                                                               String seriesTitle) {
        if (data == null) {
            throw new NoSuchElementException("Chart data does not exist");
        }
        XDDFCategoryDataSource xs = XDDFDataSourcesFactory.fromArray(categories);
        XDDFNumericalDataSource<Integer> ys = XDDFDataSourcesFactory.fromArray(values);

        XDDFChartData.Series series = data.addSeries(xs, ys);
        series.setTitle(seriesTitle, null);

        return series;
    }

    public XDDFChartData.Series createDataSeriesChartFromStringCellRange(XSSFSheet sheet, XDDFChartData data,
                                                                         int xFirstRow, int xLastRow, int xFirstCol, int xLastCol,
                                                                         int yFirstRow, int yLastRow, int yFirstCol, int yLastCol,
                                                                         String seriesTitle) {
        if (sheet == null || data == null) {
            throw new NoSuchElementException("Sheet or chart data does not exist");
        }
        XDDFCategoryDataSource xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(xFirstRow, xLastRow, xFirstCol, xLastCol));
        XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(yFirstRow, yLastRow, yFirstCol, yLastCol));

        XDDFChartData.Series series = data.addSeries(xs, ys);
        series.setTitle(seriesTitle, null);

        return series;
    }

    public void createTrendLine(CTTrendline trendLine, STTrendlineType.Enum trendLineType, STPresetLineDashVal.Enum lineDashValue,
                                int lineWidth, byte[] colors) {
        trendLine.addNewTrendlineType().setVal(trendLineType);
        trendLine.addNewSpPr().addNewLn().addNewPrstDash().setVal(lineDashValue);

        // установка ширины линии тренда
        STLineWidth stLineWidth = STLineWidth.Factory.newInstance();
        stLineWidth.setIntValue(lineWidth);
        trendLine.getSpPr().getLn().xsetW(stLineWidth);

        // установка цвета линии тренда
        trendLine.getSpPr().getLn().addNewSolidFill().addNewSrgbClr().setVal(colors);
        // не отображает формулы линии тренда
        trendLine.addNewDispEq().setVal(false);
        trendLine.addNewDispRSqr().setVal(false);
    }

    public void setChartSeriesArraySettings(CTDLbls ctdLbls, boolean showValue, boolean showCategoryName,
                                            boolean showSeriesName, boolean showPercent, boolean showLegendKey) {
        ctdLbls.addNewShowVal().setVal(showValue);
        ctdLbls.addNewShowCatName().setVal(showCategoryName);
        ctdLbls.addNewShowSerName().setVal(showSeriesName);
        ctdLbls.addNewShowPercent().setVal(showPercent);
        ctdLbls.addNewShowLegendKey().setVal(showLegendKey);
    }

    public void createCellComment(XSSFSheet sheet, int rowIndex, int cellIndex, String comment) {
        XSSFDrawing drawingPatriarch = sheet.createDrawingPatriarch();
        XSSFComment cell5Comment = drawingPatriarch.createCellComment(new XSSFClientAnchor());
        cell5Comment.setString(comment);

        XSSFRow row = sheet.getRow(rowIndex) == null ? sheet.createRow(rowIndex) : sheet.getRow(rowIndex);
        row.getCell(cellIndex).setCellComment(cell5Comment);
    }

    public XSSFCellStyle createBoldWithAllBordersStyle(XSSFWorkbook workbook, String fontName, int fontSize) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setWrapText(true);

        XSSFFont font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeight(fontSize);
        font.setBold(true);

        style.setFont(font);

        return style;
    }

    public XSSFCellStyle createRegularWithAllBordersStyle(XSSFWorkbook workbook, String fontName, int fontSize) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setWrapText(true);

        XSSFFont font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeight(fontSize);
        font.setBold(false);

        style.setFont(font);

        return style;
    }

    public XSSFCellStyle createBoldWithoutBordersStyle(XSSFWorkbook workbook, String fontName, int fontSize) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        XSSFFont font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeight(fontSize);
        font.setBold(true);

        style.setFont(font);

        return style;
    }

    public XSSFCellStyle createBoldLeftAlignmentWithoutBordersStyle(XSSFWorkbook workbook, String fontName, int fontSize) {
        XSSFCellStyle boldLeftAlignmentWithoutBordersStyle = workbook.createCellStyle();
        boldLeftAlignmentWithoutBordersStyle.setAlignment(HorizontalAlignment.LEFT);
        boldLeftAlignmentWithoutBordersStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldLeftAlignmentWithoutBordersStyle.setWrapText(true);

        XSSFFont font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeight(fontSize);
        font.setBold(true);

        boldLeftAlignmentWithoutBordersStyle.setFont(font);

        return boldLeftAlignmentWithoutBordersStyle;
    }
}
