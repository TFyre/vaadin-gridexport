package com.tfyre.vaadin.gridexport;

import com.vaadin.flow.component.grid.Grid;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * A {@link GridHolder} stub backed by plain collections, so that {@link ExcelExport} can be exercised without instantiating a Vaadin {@link Grid}.
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD {@literal <tfyre@tfyre.co.za>}
 * @param <T> the row bean type
 */
final class StubGridHolder<T> implements GridHolder<T> {

    private final List<String> columnKeys = new ArrayList<>();
    private final Map<String, String> headers = new LinkedHashMap<>();
    private final Map<String, ValueType> valueTypes = new LinkedHashMap<>();
    private final Map<String, Function<? super T, ?>> valueProviders = new LinkedHashMap<>();
    private final List<T> items;

    StubGridHolder(final List<T> items) {
        this.items = List.copyOf(items);
    }

    StubGridHolder<T> addColumn(final String key, final String headerName, final ValueType valueType,
            final Function<? super T, ?> valueProvider) {
        columnKeys.add(key);
        headers.put(key, headerName);
        valueTypes.put(key, valueType);
        valueProviders.put(key, valueProvider);
        return this;
    }

    @Override
    public Grid<T> getGrid() {
        return null;
    }

    @Override
    public List<String> getColumnKeys() {
        return Collections.unmodifiableList(columnKeys);
    }

    @Override
    public HorizontalAlignment getCellAlignment(final String columnId) {
        return valueTypes.get(columnId).isNumeric() ? HorizontalAlignment.RIGHT : HorizontalAlignment.LEFT;
    }

    @Override
    public String getColumnHeader(final String columnId) {
        return headers.get(columnId);
    }

    @Override
    public ValueType getPropertyType(final String columnId) {
        return valueTypes.get(columnId);
    }

    @Override
    public Object getPropertyValue(final T item, final String columnId) {
        return valueProviders.get(columnId).apply(item);
    }

    @Override
    public List<T> getItems() {
        return items;
    }

}
