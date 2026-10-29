package com.tfyre.vaadin.gridexport;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Tests for {@link ExcelExport}, driven by a {@link StubGridHolder} so no Vaadin {@code Grid} is required.
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD {@literal <tfyre@tfyre.co.za>}
 */
class ExcelExportTest {

    private static final Date FIRST_DATE = new Date(1_700_000_000_000L);
    private static final Date SECOND_DATE = new Date(1_710_000_000_000L);
    private static final RowData ALICE = new RowData("Alice", 3, 12.5, FIRST_DATE);
    private static final RowData BOB = new RowData("Bob", 7, 20.25, SECOND_DATE);

    private StubGridHolder<RowData> gridHolder;

    @BeforeEach
    void setUp() {
        gridHolder = new StubGridHolder<RowData>(List.of(ALICE, BOB))
                .addColumn("name", "Name", ValueType.STRING, RowData::name)
                .addColumn("quantity", "Quantity", ValueType.INTEGERBASE, RowData::quantity)
                .addColumn("price", "Price", ValueType.FLOATBASE, RowData::price)
                .addColumn("ordered", "Ordered", ValueType.DATETIME, RowData::ordered);
    }

    @Test
    void constructorAppliesDefaultSheetNameFileNameAndTotals() {
        final ExcelExport<RowData> export = new ExcelExport<>(gridHolder);

        assertEquals("Grid Export", export.getSheetName());
        assertEquals("Grid-Export.xls", export.getExportFileName());
        assertEquals("", export.getReportTitle());
        assertTrue(export.isDisplayTotals());
        assertNotNull(export.getWorkbook());
    }

    @Test
    void writesHeaderLabelsAndTypedDataCellsWithoutTitleRow() {
        final ExcelExport<RowData> export = new ExcelExport<>(gridHolder, "Data", null, null, false);
        export.getInputStream();

        final Sheet sheet = export.getWorkbook().getSheet("Data");
        assertNotNull(sheet);
        assertEquals(0, sheet.getNumMergedRegions());

        final Row header = sheet.getRow(0);
        assertEquals("Name", header.getCell(0).getStringCellValue());
        assertEquals("Quantity", header.getCell(1).getStringCellValue());
        assertEquals("Price", header.getCell(2).getStringCellValue());
        assertEquals("Ordered", header.getCell(3).getStringCellValue());

        final Row firstDataRow = sheet.getRow(1);
        assertEquals(CellType.STRING, firstDataRow.getCell(0).getCellType());
        assertEquals("Alice", firstDataRow.getCell(0).getStringCellValue());
        assertEquals(CellType.NUMERIC, firstDataRow.getCell(1).getCellType());
        assertEquals(3.0, firstDataRow.getCell(1).getNumericCellValue());
        assertEquals(CellType.NUMERIC, firstDataRow.getCell(2).getCellType());
        assertEquals(12.5, firstDataRow.getCell(2).getNumericCellValue());
        assertEquals(CellType.NUMERIC, firstDataRow.getCell(3).getCellType());
        assertEquals(FIRST_DATE, firstDataRow.getCell(3).getDateCellValue());

        final Row secondDataRow = sheet.getRow(2);
        assertEquals("Bob", secondDataRow.getCell(0).getStringCellValue());
        assertEquals(7.0, secondDataRow.getCell(1).getNumericCellValue());
        assertEquals(20.25, secondDataRow.getCell(2).getNumericCellValue());
        assertEquals(SECOND_DATE, secondDataRow.getCell(3).getDateCellValue());
    }

    @Test
    void appliesColumnAlignmentToDataCells() {
        final ExcelExport<RowData> export = new ExcelExport<>(gridHolder, "Data", null, null, false);
        export.getInputStream();

        final Row dataRow = export.getWorkbook().getSheet("Data").getRow(1);
        assertEquals(HorizontalAlignment.LEFT, dataRow.getCell(0).getCellStyle().getAlignment());
        assertEquals(HorizontalAlignment.RIGHT, dataRow.getCell(1).getCellStyle().getAlignment());
        assertEquals(HorizontalAlignment.RIGHT, dataRow.getCell(2).getCellStyle().getAlignment());
        assertEquals(HorizontalAlignment.LEFT, dataRow.getCell(3).getCellStyle().getAlignment());
    }

    @Test
    void appliesDefaultNumberFormatsByValueType() {
        final ExcelExport<RowData> export = new ExcelExport<>(gridHolder, "Data", null, null, false);
        export.getInputStream();

        final Row dataRow = export.getWorkbook().getSheet("Data").getRow(1);
        assertEquals("0", dataRow.getCell(1).getCellStyle().getDataFormatString());
        assertEquals("0.00", dataRow.getCell(2).getCellStyle().getDataFormatString());
        assertEquals("mm/dd/yyyy", dataRow.getCell(3).getCellStyle().getDataFormatString());
    }

    @Test
    void appliesPerColumnExcelFormat() {
        final ExcelExport<RowData> export = new ExcelExport<>(gridHolder, "Data", null, null, false);
        export.setExcelFormatOfProperty("price", "#,##0.000");
        export.getInputStream();

        final Row dataRow = export.getWorkbook().getSheet("Data").getRow(1);
        assertEquals("0", dataRow.getCell(1).getCellStyle().getDataFormatString());
        assertEquals("#,##0.000", dataRow.getCell(2).getCellStyle().getDataFormatString());
    }

    @Test
    void writesTotalsRowWithSumFormulasAndTotalLabel() {
        final ExcelExport<RowData> export = new ExcelExport<>(gridHolder);
        export.getInputStream();

        final Sheet sheet = export.getWorkbook().getSheet("Grid Export");
        assertEquals(3, sheet.getLastRowNum());

        final Row totalsRow = sheet.getRow(3);
        assertSame(totalsRow, export.getTotalsRow());
        assertEquals("Total", totalsRow.getCell(0).getStringCellValue());

        final Cell quantityTotal = totalsRow.getCell(1);
        assertEquals(CellType.FORMULA, quantityTotal.getCellType());
        assertEquals("SUM(B1:B3)", quantityTotal.getCellFormula());
        assertEquals(10.0, quantityTotal.getNumericCellValue());

        final Cell priceTotal = totalsRow.getCell(2);
        assertEquals(CellType.FORMULA, priceTotal.getCellType());
        assertEquals("SUM(C1:C3)", priceTotal.getCellFormula());
        assertEquals(32.75, priceTotal.getNumericCellValue());

        assertEquals(CellType.BLANK, totalsRow.getCell(3).getCellType());
    }

    @Test
    void omitsTotalsRowWhenDisabled() {
        final ExcelExport<RowData> export = new ExcelExport<>(gridHolder, "Data", null, null, false);
        export.getInputStream();

        final Sheet sheet = export.getWorkbook().getSheet("Data");
        assertEquals(2, sheet.getLastRowNum());
        assertNull(sheet.getRow(3));
        assertNull(export.getTotalsRow());
    }

    @Test
    void writesTitleRowAndShiftsHeaderDataAndTotalsDown() {
        final ExcelExport<RowData> export = new ExcelExport<>(gridHolder, "Report", "Quarterly Report");
        export.getInputStream();

        final Workbook workbook = export.getWorkbook();
        assertEquals("Report", workbook.getSheetName(0));
        final Sheet sheet = workbook.getSheet("Report");

        final Row titleRow = sheet.getRow(0);
        assertEquals("Quarterly Report", titleRow.getCell(0).getStringCellValue());
        assertEquals(1, sheet.getNumMergedRegions());
        final CellRangeAddress mergedRegion = sheet.getMergedRegion(0);
        assertEquals(0, mergedRegion.getFirstRow());
        assertEquals(0, mergedRegion.getLastRow());
        assertEquals(0, mergedRegion.getFirstColumn());
        assertEquals(gridHolder.getColumnKeys().size() - 1, mergedRegion.getLastColumn());

        assertEquals("Name", sheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals("Alice", sheet.getRow(2).getCell(0).getStringCellValue());
        assertEquals("Bob", sheet.getRow(3).getCell(0).getStringCellValue());
        assertEquals("Total", sheet.getRow(4).getCell(0).getStringCellValue());
        assertEquals("SUM(B2:B4)", sheet.getRow(4).getCell(1).getCellFormula());
    }

    @Test
    void singleGetInputStreamCallWritesExactlyOneHeaderAndDataBlock() throws IOException {
        final ExcelExport<RowData> export = new ExcelExport<>(gridHolder, "Data", null, null, false);

        final byte[] bytes;
        try (InputStream inputStream = export.getInputStream()) {
            bytes = inputStream.readAllBytes();
        }
        assertTrue(bytes.length > 0);

        try (Workbook parsed = new HSSFWorkbook(new ByteArrayInputStream(bytes))) {
            final Sheet sheet = parsed.getSheet("Data");
            assertNotNull(sheet);
            assertEquals(3, sheet.getPhysicalNumberOfRows());
            assertEquals(2, sheet.getLastRowNum());
            assertEquals(1, countStringCells(sheet, 0, "Name"));
            assertEquals(1, countStringCells(sheet, 0, "Alice"));
            assertEquals(1, countStringCells(sheet, 0, "Bob"));
        }
    }

    private static long countStringCells(final Sheet sheet, final int column, final String value) {
        long count = 0;
        for (final Row row : sheet) {
            final Cell cell = row.getCell(column);
            if (cell != null && cell.getCellType() == CellType.STRING && value.equals(cell.getStringCellValue())) {
                count++;
            }
        }
        return count;
    }

    private record RowData(String name, int quantity, double price, Date ordered) {
    }

}
