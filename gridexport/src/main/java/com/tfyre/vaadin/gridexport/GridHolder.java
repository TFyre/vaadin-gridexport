package com.tfyre.vaadin.gridexport;

import com.vaadin.data.ValueProvider;
import com.vaadin.ui.Grid;
import java.io.Serializable;
import java.util.Collection;

import com.vaadin.ui.UI;
import java.util.List;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD <tfyre@tfyre.co.za>
 * @param <T> the grid bean type
 */
public interface GridHolder<T> extends Serializable {

    Grid<T> getGrid();

    void hideColumn(final String propId);

    List<String> getPropIds();

    boolean isHierarchical();

    void setHierarchical(final boolean hierarchical);

    HorizontalAlignment getCellAlignment(String propId);

    // grid delegated methods
    boolean isColumnCollapsed(String propertyId);

    UI getUI();

    String getColumnHeader(String propertyId);

    abstract Class<?> getPropertyType(String propId);

    abstract Object getPropertyValue(T itemId, String propId);

    abstract void setColumnValueProvider(String propId, ValueProvider<T, ?> valueProvider);

    abstract Collection<T> getItemIds();

    abstract Collection<T> getRootItemIds();
}
