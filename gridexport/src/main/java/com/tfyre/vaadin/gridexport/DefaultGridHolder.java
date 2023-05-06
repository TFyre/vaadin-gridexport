package com.tfyre.vaadin.gridexport;

import com.vaadin.flow.component.grid.Grid;
import com.vaadin.flow.function.ValueProvider;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD {@literal <tfyre@tfyre.co.za>}
 * @param <T> the grid bean type
 */
public class DefaultGridHolder<T> implements GridHolder<T> {

    protected final HorizontalAlignment defaultAlignment = HorizontalAlignment.LEFT;

    private final Grid<T> grid;
    private final Map<String, ColumnDetail<T>> columns = new HashMap<>();
    private final List<String> columnKeys = new ArrayList<>();

    public DefaultGridHolder(final Grid<T> grid) {
        this.grid = grid;
    }

    @Override
    public Grid<T> getGrid() {
        return grid;
    }

    public DefaultGridHolder<T> clearColumns() {
        columnKeys.clear();
        columns.clear();
        return this;
    }

    public DefaultGridHolder<T> addColumn(final String key, final Grid.Column<T> column, final ValueProvider<T, ?> valueProvider, final String headerName,
            final ValueType valueType) {
        columnKeys.add(key);
        columns.put(key, new ColumnDetail<>(key, column, valueProvider, headerName, valueType));
        return this;
    }

    @Override
    public List<String> getColumnKeys() {
        return Collections.unmodifiableList(columnKeys);
    }

    @Override
    public HorizontalAlignment getCellAlignment(final String columnId) {
        return getPropertyType(columnId).isNumeric() ? HorizontalAlignment.RIGHT : defaultAlignment;
    }

    @Override
    public String getColumnHeader(final String columnId) {
        return columns.get(columnId).headerName();
    }

    protected Grid.Column<T> getColumn(final String columnId) {
        return columns.get(columnId).column();
    }

    @Override
    public ValueType getPropertyType(final String columnId) {
        return columns.get(columnId).valueType();
    }

    @Override
    public Object getPropertyValue(final T item, final String columnId) {
        return columns.get(columnId).valueProvider().apply(item);
    }

    @Override
    public List<T> getItems() {
        return grid.getDataProvider().fetch(grid.getDataCommunicator().buildQuery(0, Integer.MAX_VALUE)).toList();
    }

    private record ColumnDetail<S>(String key, Grid.
    Column<S> column, ValueProvider<S,?>valueProvider,
    String headerName, ValueType valueType)
    {

    }

}
