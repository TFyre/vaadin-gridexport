package com.tfyre.vaadin.gridexport;

import com.vaadin.flow.component.grid.Grid;
import java.util.List;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD {@literal <tfyre@tfyre.co.za>}
 * @param <T> the grid bean type
 */
public interface GridHolder<T>  {

    Grid<T> getGrid();

    List<String> getColumnKeys();

    HorizontalAlignment getCellAlignment(final String columnId);

    String getColumnHeader(final String columnId);

    abstract ValueType getPropertyType(final String columnId);

    abstract Object getPropertyValue(T itemId, final String columnId);

    abstract List<T> getItems();

}
