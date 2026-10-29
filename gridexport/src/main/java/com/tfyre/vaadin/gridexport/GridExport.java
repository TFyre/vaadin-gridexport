package com.tfyre.vaadin.gridexport;

import java.io.InputStream;

/**
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD {@literal <tfyre@tfyre.co.za>}
 * @param <T> the grid bean type
 */
public interface GridExport<T> {

    GridHolder<T> getGridHolder();

    void convertGrid();

    InputStream getInputStream();

}
