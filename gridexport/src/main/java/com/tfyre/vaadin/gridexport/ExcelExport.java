package com.tfyre.vaadin.gridexport;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;

import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD {@literal <tfyre@tfyre.co.za>}
 * @param <T> the grid bean type
 */
public class ExcelExport<T> implements GridExport<T> {

    private static final Logger LOGGER = Logger.getLogger(ExcelExport.class.getName());

    private String sheetName;
    private String reportTitle;
    private String exportFileName;
    private boolean displayTotals;
    private boolean rowHeaders = false;
    private final Workbook workbook;
    private Sheet sheet;
    private CreationHelper createHelper;
    private DataFormat dataFormat;
    private CellStyle dateCellStyle, doubleCellStyle, integerCellStyle, totalsDoubleCellStyle,
            totalsIntegerCellStyle, columnHeaderCellStyle, titleCellStyle;
    private Short dateDataFormat, doubleDataFormat, integerDataFormat;
    private final Map<Short, CellStyle> dataFormatCellStylesMap = new HashMap<>();
    private CellStyle rowHeaderCellStyle = null;
    private Row titleRow, headerRow, totalsRow;
    private final Map<Object, String> propertyExcelFormatMap = new HashMap<>();
    private final GridHolder<T> gridHolder;

    public ExcelExport(final GridHolder<T> gridHolder) {
        this(gridHolder, null);
    }

    public ExcelExport(final GridHolder<T> gridHolder, final String sheetName) {
        this(gridHolder, sheetName, null);
    }

    public ExcelExport(final GridHolder<T> gridHolder, final String sheetName, final String reportTitle) {
        this(gridHolder, sheetName, reportTitle, null);
    }

    public ExcelExport(final GridHolder<T> gridHolder, final String sheetName, final String reportTitle,
            final String exportFileName) {
        this(gridHolder, sheetName, reportTitle, exportFileName, true);
    }

    public ExcelExport(final GridHolder<T> gridHolder, final String sheetName, final String reportTitle,
            final String exportFileName, final boolean hasTotalsRow) {
        this(gridHolder, new HSSFWorkbook(), sheetName, reportTitle, exportFileName, hasTotalsRow);
    }

    public ExcelExport(final GridHolder<T> gridHolder, final Workbook wkbk, final String shtName,
            final String rptTitle, final String xptFileName, final boolean hasTotalsRow) {
        this.gridHolder = gridHolder;
        this.workbook = wkbk;
        init(shtName, rptTitle, xptFileName, hasTotalsRow);
    }

    @Override
    public GridHolder<T> getGridHolder() {
        return gridHolder;
    }

    private void init(final String shtName, final String rptTitle, final String xptFileName,
            final boolean hasTotalsRow) {
        if ((null == shtName) || ("".equals(shtName))) {
            this.sheetName = "Grid Export";
        } else {
            this.sheetName = shtName;
        }
        if (null == rptTitle) {
            this.reportTitle = "";
        } else {
            this.reportTitle = rptTitle;
        }
        if ((null == xptFileName) || ("".equals(xptFileName))) {
            this.exportFileName = "Grid-Export.xls";
        } else {
            this.exportFileName = xptFileName;
        }
        this.displayTotals = hasTotalsRow;

        this.sheet = this.workbook.createSheet(this.sheetName);
        this.createHelper = this.workbook.getCreationHelper();
        this.dataFormat = this.workbook.createDataFormat();
        this.dateDataFormat = defaultDateDataFormat(this.workbook);
        this.doubleDataFormat = defaultDoubleDataFormat(this.workbook);
        this.integerDataFormat = defaultIntegerDataFormat(this.workbook);

        this.doubleCellStyle = defaultDataCellStyle(this.workbook);
        this.doubleCellStyle.setDataFormat(doubleDataFormat);
        this.dataFormatCellStylesMap.put(doubleDataFormat, doubleCellStyle);

        this.integerCellStyle = defaultDataCellStyle(this.workbook);
        this.integerCellStyle.setDataFormat(integerDataFormat);
        this.dataFormatCellStylesMap.put(integerDataFormat, integerCellStyle);

        this.dateCellStyle = defaultDataCellStyle(this.workbook);
        this.dateCellStyle.setDataFormat(this.dateDataFormat);
        this.dataFormatCellStylesMap.put(this.dateDataFormat, this.dateCellStyle);

        this.totalsDoubleCellStyle = defaultTotalsDoubleCellStyle(this.workbook);
        this.totalsIntegerCellStyle = defaultTotalsIntegerCellStyle(this.workbook);
        this.columnHeaderCellStyle = defaultHeaderCellStyle(this.workbook);
        this.titleCellStyle = defaultTitleCellStyle(this.workbook);
    }

    /**
     * Creates the workbook containing the exported grid data, without exporting it to the user.
     */
    @Override
    public void convertGrid() {
        final int startRow;
        // initial setup
        initialSheetSetup();

        // add title row
        startRow = addTitleRow();
        int row = startRow;

        // add header row
        addHeaderRow(row);
        row++;

        // add data rows
        row = addDataRows(sheet, row);

        // add totals row
        if (displayTotals) {
            addTotalsRow(row, startRow);
        }

        // final sheet format before export
        finalSheetFormat();
    }

    @Override
    public InputStream getInputStream() {
        convertGrid();

        final ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try {
            workbook.write(bos);
        } catch (IOException ex) {
            Logger.getLogger(ExcelExport.class.getName()).log(Level.SEVERE, null, ex);
        }
        return new ByteArrayInputStream(bos.toByteArray());
    }

    /**
     * Initial sheet setup. Override this method to specifically change initial, sheet-wide, settings.
     */
    protected void initialSheetSetup() {
        final PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);
    }

    /**
     * Adds the title row. Override this method to change title-related aspects of the workbook. Alternately, the title Row Object is accessible via
     * getTitleRow() after report creation. To change title text use setReportTitle(). To change title CellStyle use setTitleStyle().
     *
     * @return the int
     */
    protected int addTitleRow() {
        if ((null == reportTitle) || ("".equals(reportTitle))) {
            return 0;
        }
        titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        final Cell titleCell;
        final CellRangeAddress cra;
        final int size = gridHolder.getColumnKeys().size();
        if (rowHeaders) {
            titleCell = titleRow.createCell(1);
            cra = new CellRangeAddress(0, 0, 1, size - 1);
            sheet.addMergedRegion(cra);
        } else {
            titleCell = titleRow.createCell(0);
            cra = new CellRangeAddress(0, 0, 0, size - 1);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, size - 1));
        }
        titleCell.setCellValue(reportTitle);
        titleCell.setCellStyle(titleCellStyle);
        // cell borders don't work on merged ranges so, if there are borders
        // we apply them to the merged range here.
        if (titleCellStyle.getBorderLeft() != BorderStyle.NONE) {
            RegionUtil.setBorderLeft(titleCellStyle.getBorderLeft(), cra, sheet);
        }
        if (titleCellStyle.getBorderRight() != BorderStyle.NONE) {
            RegionUtil.setBorderRight(titleCellStyle.getBorderRight(), cra, sheet);
        }
        if (titleCellStyle.getBorderTop() != BorderStyle.NONE) {
            RegionUtil.setBorderTop(titleCellStyle.getBorderTop(), cra, sheet);
        }
        if (titleCellStyle.getBorderBottom() != BorderStyle.NONE) {
            RegionUtil.setBorderBottom(titleCellStyle.getBorderBottom(), cra, sheet);
        }
        return 1;
    }

    /**
     * Adds the header row. Override this method to change header-row-related aspects of the workbook. Alternately, the header Row Object is accessible via
     * getHeaderRow() after report creation. To change header CellStyle, though, use setHeaderStyle().
     *
     * @param row the row
     */
    protected void addHeaderRow(final int row) {
        headerRow = sheet.createRow(row);
        Cell headerCell;
        headerRow.setHeightInPoints(40);
        int col = 0;
        for (final String columnKey : gridHolder.getColumnKeys()) {
            headerCell = headerRow.createCell(col);
            headerCell.setCellValue(createHelper.createRichTextString(getGridHolder().getColumnHeader(columnKey)));
            headerCell.setCellStyle(getColumnHeaderStyle(row, col));

            final HorizontalAlignment poiAlignment = getGridHolder().getCellAlignment(columnKey);
            CellUtil.setAlignment(headerCell, poiAlignment);
            col++;
        }
    }

    /**
     * This method is called by addTotalsRow() to determine what CellStyle to use. By default we just return totalsCellStyle which is either set to the default
     * totals style, or can be overriden by the user using setTotalsStyle(). However, if the user wants to have different total items have different styles,
     * then this method should be overriden. The parameters passed in are all potentially relevant items that may be used to determine what formatting to
     * return, that are not accessible globally.
     *
     * @param row the row
     * @param col the current column
     * @return the header style
     */
    protected CellStyle getColumnHeaderStyle(final int row, final int col) {
        if ((rowHeaders) && (col == 0)) {
            return titleCellStyle;
        }
        return columnHeaderCellStyle;
    }

    /**
     * this method adds row items for non-Hierarchical Containers. Override this method to make any changes. To change the CellStyle used for all Grid data use
     * setDataStyle(). For different data cells to have different CellStyles, override getDataStyle().
     *
     * @param sheetToAddTo the sheet to add to
     * @param row the row
     * @return the int
     */
    protected int addDataRows(final Sheet sheetToAddTo, final int row) {
        final AtomicInteger localRow = new AtomicInteger(row);
        getGridHolder().getItems().forEach(itemId -> {
            addDataRow(sheetToAddTo, itemId, localRow.getAndIncrement());
        });
        return localRow.get();
    }

    /**
     * This method is ultimately used by either addDataRows() to actually add the data to the Sheet.
     *
     * @param sheetToAddTo the sheet to add to
     * @param rootItemId the root item id
     * @param row the row
     */
    protected void addDataRow(final Sheet sheetToAddTo, final T rootItemId, final int row) {
        final Row sheetRow = sheetToAddTo.createRow(row);
        int col = 0;
        for (final String columnKey : gridHolder.getColumnKeys()) {
            setupCell(sheetRow.createCell(col),
                    getGridHolder().getPropertyValue(rootItemId, columnKey),
                    columnKey, row, col);
            col++;
        }
    }

    protected void setupCell(Cell sheetCell, Object value, String columnKey, int row, int col) {
        sheetCell.setCellStyle(getCellStyle(columnKey, row, col, false));
        final HorizontalAlignment poiAlignment = getGridHolder().getCellAlignment(columnKey);
        CellUtil.setAlignment(sheetCell, poiAlignment);
        setCellValue(sheetCell, value, columnKey);
    }

    protected void setCellValue(Cell sheetCell, Object value, String columnKey) {
        if (null != value) {
            final ValueType valueType = getGridHolder().getPropertyType(columnKey);
            if (valueType == ValueType.DATETIME) {
                sheetCell.setCellValue((Date) value);
            } else if (valueType.isNumeric()) {
                try {
                    // parse all numbers as double, the format will determine how they appear
                    final Double d = Double.parseDouble(value.toString());
                    sheetCell.setCellValue(d);
                } catch (final NumberFormatException nfe) {
                    LOGGER.warning("NumberFormatException parsing a numeric value: " + nfe);
                    sheetCell.setCellValue(createHelper.createRichTextString(value.toString()));
                }
            } else {
                sheetCell.setCellValue(createHelper.createRichTextString(value.toString()));
            }
        }
    }

    public void setExcelFormatOfProperty(final Object propertyId, final String excelFormat) {
        if (this.propertyExcelFormatMap.containsKey(propertyId)) {
            this.propertyExcelFormatMap.remove(propertyId);
        }
        this.propertyExcelFormatMap.put(propertyId.toString(), excelFormat);
    }

    /**
     * This method is called by addDataRow() to determine what CellStyle to use. By default we just return dataStyle which is either set to the default data
     * style, or can be overriden by the user using setDataStyle(). However, if the user wants to have different data items have different styles, then this
     * method should be overriden. The parameters passed in are all potentially relevant items that may be used to determine what formatting to return, that are
     * not accessible globally.
     *
     * @param columnKey the property id
     * @param row the row
     * @param col the col
     * @param totalsRow show totals row
     * @return the data style
     */
    protected CellStyle getCellStyle(final String columnKey, final int row, final int col, final boolean totalsRow) {
        // get the basic style for the type of cell (i.e. data, header, total)
        if ((rowHeaders) && (col == 0)) {
            if (null == rowHeaderCellStyle) {
                return columnHeaderCellStyle;
            }
            return rowHeaderCellStyle;
        }
        final ValueType propType = getGridHolder().getPropertyType(columnKey);
        if (totalsRow) {
            if (this.propertyExcelFormatMap.containsKey(columnKey)) {
                final short df = dataFormat.getFormat(propertyExcelFormatMap.get(columnKey));
                final CellStyle customTotalStyle = workbook.createCellStyle();
                customTotalStyle.cloneStyleFrom(totalsDoubleCellStyle);
                customTotalStyle.setDataFormat(df);
                return customTotalStyle;
            }
            if (propType == ValueType.INTEGERBASE) {
                return totalsIntegerCellStyle;
            }
            return totalsDoubleCellStyle;
        }
        // Check if the user has over-ridden that data format of this property
        if (this.propertyExcelFormatMap.containsKey(columnKey)) {
            final short df = dataFormat.getFormat(propertyExcelFormatMap.get(columnKey));
            if (dataFormatCellStylesMap.containsKey(df)) {
                return dataFormatCellStylesMap.get(df);
            }
            // if it hasn't already been created for re-use, we create a cell style and override the data format
            // For data cells, each data format corresponds to a single complete cell style
            final CellStyle retStyle = workbook.createCellStyle();
            retStyle.cloneStyleFrom(dataFormatCellStylesMap.get(doubleDataFormat));
            retStyle.setDataFormat(df);
            dataFormatCellStylesMap.put(df, retStyle);
            return retStyle;
        }
        // if not over-ridden, use the overall setting
        return switch (propType) {
            case INTEGERBASE -> dataFormatCellStylesMap.get(integerDataFormat);
            case DATETIME -> dataFormatCellStylesMap.get(dateDataFormat);
            default -> dataFormatCellStylesMap.get(doubleDataFormat);
        };
    }

    /**
     * Adds the totals row to the report. Override this method to make any changes. Alternately, the totals Row Object is accessible via getTotalsRow() after
     * report creation. To change the CellStyle used for the totals row, use setFormulaStyle. For different totals cells to have different CellStyles, override
     * getTotalsStyle().
     *
     * @param currentRow the current row
     * @param startRow the start row
     */
    protected void addTotalsRow(final int currentRow, final int startRow) {
        totalsRow = sheet.createRow(currentRow);
        totalsRow.setHeightInPoints(30);
        int col = 0;
        for (final String columnKey : gridHolder.getColumnKeys()) {
            setupTotalCell(totalsRow.createCell(col), columnKey, currentRow, startRow, col);
            col++;
        }
    }

    protected void setupTotalCell(Cell cell, final String columnKey, final int currentRow, final int startRow, int col) {
        cell.setCellStyle(getCellStyle(columnKey, startRow, col, true));
        final HorizontalAlignment poiAlignment = getGridHolder().getCellAlignment(columnKey);
        CellUtil.setAlignment(cell, poiAlignment);
        if (getGridHolder().getPropertyType(columnKey).isNumeric()) {
            CellRangeAddress cra = new CellRangeAddress(startRow, currentRow - 1, col, col);
            cell.setCellFormula("SUM(" + cra.formatAsString() + ")");
        } else {
            if (0 == col) {
                cell.setCellValue(createHelper.createRichTextString("Total"));
            }
        }
    }

    /**
     * Final formatting of the sheet upon completion of writing the data. For example, we can only size the column widths once the data is in the report and the
     * sheet knows how wide the data is.
     */
    protected void finalSheetFormat() {
        final FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
        for (int col = 0; col < gridHolder.getColumnKeys().size(); col++) {
            sheet.autoSizeColumn(col);
        }
    }

    /**
     * Returns the default title style. Obtained from: http://svn.apache.org/repos/asf/poi /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultTitleCellStyle(final Workbook wb) {
        CellStyle style;
        final Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 18);
        titleFont.setBold(true);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(titleFont);
        return style;
    }

    /**
     * Returns the default header style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultHeaderCellStyle(final Workbook wb) {
        CellStyle style;
        final Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short) 11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(monthFont);
        style.setWrapText(true);
        return style;
    }

    /**
     * Returns the default data cell style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultDataCellStyle(final Workbook wb) {
        final CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setDataFormat(doubleDataFormat);
        return style;
    }

    /**
     * Returns the default totals row style for Double data. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultTotalsDoubleCellStyle(final Workbook wb) {
        final CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(doubleDataFormat);
        return style;
    }

    /**
     * Returns the default totals row style for Integer data. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultTotalsIntegerCellStyle(final Workbook wb) {
        final CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(integerDataFormat);
        return style;
    }

    protected short defaultDoubleDataFormat(final Workbook wb) {
        return createHelper.createDataFormat().getFormat("0.00");
    }

    protected short defaultIntegerDataFormat(final Workbook wb) {
        return createHelper.createDataFormat().getFormat("0");
    }

    protected short defaultDateDataFormat(final Workbook wb) {
        return createHelper.createDataFormat().getFormat("mm/dd/yyyy");
    }

    public void setDoubleDataFormat(final String excelDoubleFormat) {
        CellStyle prevDoubleDataStyle = null;
        if (dataFormatCellStylesMap.containsKey(doubleDataFormat)) {
            prevDoubleDataStyle = dataFormatCellStylesMap.get(doubleDataFormat);
            dataFormatCellStylesMap.remove(doubleDataFormat);
        }
        doubleDataFormat = createHelper.createDataFormat().getFormat(excelDoubleFormat);
        if (null != prevDoubleDataStyle) {
            doubleCellStyle = prevDoubleDataStyle;
            doubleCellStyle.setDataFormat(doubleDataFormat);
            dataFormatCellStylesMap.put(doubleDataFormat, doubleCellStyle);
        }
    }

    public void setIntegerDataFormat(final String excelIntegerFormat) {
        CellStyle prevIntegerDataStyle = null;
        if (dataFormatCellStylesMap.containsKey(integerDataFormat)) {
            prevIntegerDataStyle = dataFormatCellStylesMap.get(integerDataFormat);
            dataFormatCellStylesMap.remove(integerDataFormat);
        }
        integerDataFormat = createHelper.createDataFormat().getFormat(excelIntegerFormat);
        if (null != prevIntegerDataStyle) {
            integerCellStyle = prevIntegerDataStyle;
            integerCellStyle.setDataFormat(integerDataFormat);
            dataFormatCellStylesMap.put(integerDataFormat, integerCellStyle);
        }
    }

    public void setDateDataFormat(final String excelDateFormat) {
        CellStyle prevDateDataStyle = null;
        if (dataFormatCellStylesMap.containsKey(dateDataFormat)) {
            prevDateDataStyle = dataFormatCellStylesMap.get(dateDataFormat);
            dataFormatCellStylesMap.remove(dateDataFormat);
        }
        dateDataFormat = createHelper.createDataFormat().getFormat(excelDateFormat);
        if (null != prevDateDataStyle) {
            dateCellStyle = prevDateDataStyle;
            dateCellStyle.setDataFormat(dateDataFormat);
            dataFormatCellStylesMap.put(dateDataFormat, dateCellStyle);
        }
    }

    /**
     * Gets the workbook.
     *
     * @return the workbook
     */
    public Workbook getWorkbook() {
        return this.workbook;
    }

    /**
     * Gets the sheet name.
     *
     * @return the sheet name
     */
    public String getSheetName() {
        return this.sheetName;
    }

    /**
     * Gets the report title.
     *
     * @return the report title
     */
    public String getReportTitle() {
        return this.reportTitle;
    }

    /**
     * Gets the export file name.
     *
     * @return the export file name
     */
    public String getExportFileName() {
        return this.exportFileName;
    }

    /**
     * Gets the cell style used for report data..
     *
     * @return the cell style
     */
    public CellStyle getDoubleDataStyle() {
        return this.doubleCellStyle;
    }

    /**
     * Gets the cell style used for report data..
     *
     * @return the cell style
     */
    public CellStyle getIntegerDataStyle() {
        return this.integerCellStyle;
    }

    public CellStyle getDateDataStyle() {
        return this.dateCellStyle;
    }

    /**
     * Gets the cell style used for the report headers.
     *
     * @return the column header style
     */
    public CellStyle getColumnHeaderStyle() {
        return this.columnHeaderCellStyle;
    }

    /**
     * Gets the cell title used for the report title.
     *
     * @return the title style
     */
    public CellStyle getTitleStyle() {
        return this.titleCellStyle;
    }

    /**
     * Sets the text used for the report title.
     *
     * @param reportTitle the new report title
     */
    public void setReportTitle(final String reportTitle) {
        this.reportTitle = reportTitle;
    }

    /**
     * Sets the export file name.
     *
     * @param exportFileName the new export file name
     */
    public void setExportFileName(final String exportFileName) {
        this.exportFileName = exportFileName;
    }

    /**
     * Sets the cell style used for report data.
     *
     * @param doubleDataStyle the new data style
     */
    public void setDoubleDataStyle(final CellStyle doubleDataStyle) {
        this.doubleCellStyle = doubleDataStyle;
    }

    /**
     * Sets the cell style used for report data.
     *
     * @param integerDataStyle the new data style
     */
    public void setIntegerDataStyle(final CellStyle integerDataStyle) {
        this.integerCellStyle = integerDataStyle;
    }

    /**
     * Sets the cell style used for report data.
     *
     * @param dateDataStyle the new data style
     */
    public void setDateDataStyle(final CellStyle dateDataStyle) {
        this.dateCellStyle = dateDataStyle;
    }

    /**
     * Sets the cell style used for the report headers.
     *
     * @param columnHeaderStyle CellStyle
     */
    public void setColumnHeaderStyle(final CellStyle columnHeaderStyle) {
        this.columnHeaderCellStyle = columnHeaderStyle;
    }

    /**
     * Sets the cell style used for the report title.
     *
     * @param titleStyle the new title style
     */
    public void setTitleStyle(final CellStyle titleStyle) {
        this.titleCellStyle = titleStyle;
    }

    /**
     * Gets the title row.
     *
     * @return the title row
     */
    public Row getTitleRow() {
        return this.titleRow;
    }

    /**
     * Gets the header row.
     *
     * @return the header row
     */
    public Row getHeaderRow() {
        return this.headerRow;
    }

    /**
     * Gets the totals row.
     *
     * @return the totals row
     */
    public Row getTotalsRow() {
        return this.totalsRow;
    }

    /**
     * Gets the cell style used for the totals row.
     *
     * @return the totals style
     */
    public CellStyle getTotalsDoubleStyle() {
        return this.totalsDoubleCellStyle;
    }

    /**
     * Sets the cell style used for the totals row.
     *
     * @param totalsDoubleStyle the new totals style
     */
    public void setTotalsDoubleStyle(final CellStyle totalsDoubleStyle) {
        this.totalsDoubleCellStyle = totalsDoubleStyle;
    }

    /**
     * Gets the cell style used for the totals row.
     *
     * @return the totals style
     */
    public CellStyle getTotalsIntegerStyle() {
        return this.totalsIntegerCellStyle;
    }

    /**
     * Sets the cell style used for the totals row.
     *
     * @param totalsIntegerStyle the new totals style
     */
    public void setTotalsIntegerStyle(final CellStyle totalsIntegerStyle) {
        this.totalsIntegerCellStyle = totalsIntegerStyle;
    }

    /**
     * Flag indicating whether a totals row will be added to the report or not.
     *
     * @return true, if totals row will be added
     */
    public boolean isDisplayTotals() {
        return this.displayTotals;
    }

    /**
     * Sets the flag indicating whether a totals row will be added to the report or not.
     *
     * @param displayTotals boolean
     */
    public void setDisplayTotals(final boolean displayTotals) {
        this.displayTotals = displayTotals;
    }

    /**
     * See value of flag indicating whether the first column should be treated as row headers.
     *
     * @return boolean
     */
    public boolean hasRowHeaders() {
        return this.rowHeaders;
    }

    /**
     * Method getRowHeaderStyle.
     *
     * @return CellStyle
     */
    public CellStyle getRowHeaderStyle() {
        return this.rowHeaderCellStyle;
    }

    /**
     * Set value of flag indicating whether the first column should be treated as row headers.
     *
     * @param rowHeaders boolean
     */
    public void setRowHeaders(final boolean rowHeaders) {
        this.rowHeaders = rowHeaders;
    }

    /**
     * Method setRowHeaderStyle.
     *
     * @param rowHeaderStyle CellStyle
     */
    public void setRowHeaderStyle(final CellStyle rowHeaderStyle) {
        this.rowHeaderCellStyle = rowHeaderStyle;
    }

}
