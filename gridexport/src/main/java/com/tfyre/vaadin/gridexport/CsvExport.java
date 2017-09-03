package com.tfyre.vaadin.gridexport;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD <tfyre@tfyre.co.za>
 * @param <T> the grid bean type
 */
public class CsvExport<T> extends ExcelExport<T> {

    private static final long serialVersionUID = VersionHelper.serialVersionUID;
    private static final Logger LOG = Logger.getLogger(CsvExport.class.getName());

    public CsvExport(final GridHolder<T> gridHolder) {
        super(gridHolder);
    }

    public CsvExport(final GridHolder<T> gridHolder, final String sheetName) {
        super(gridHolder, sheetName);
    }

    public CsvExport(final GridHolder<T> gridHolder, final String sheetName, final String reportTitle) {
        super(gridHolder, sheetName, reportTitle);
    }

    public CsvExport(final GridHolder<T> gridHolder, final String sheetName, final String reportTitle,
            final String exportFileName) {
        super(gridHolder, sheetName, reportTitle, exportFileName);
    }

    public CsvExport(final GridHolder<T> gridHolder, final String sheetName, final String reportTitle,
            final String exportFileName, final boolean hasTotalsRow) {
        super(gridHolder, sheetName, reportTitle, exportFileName, hasTotalsRow);
    }

    @Override
    public InputStream getInputStream() {
        try {
            final POIFSFileSystem fs = new POIFSFileSystem(super.getInputStream());
            final ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (final PrintStream p = new PrintStream(bos)) {
                final XLS2CSVmra xls2csv = new XLS2CSVmra(fs, p, -1);
                xls2csv.process();
            }
            return new ByteArrayInputStream(bos.toByteArray());
        } catch (IOException ex) {
            LOG.log(Level.SEVERE, null, ex);
            return new ByteArrayInputStream(new byte[]{});
        }
    }

}
