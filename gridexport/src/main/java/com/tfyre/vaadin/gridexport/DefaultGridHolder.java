package com.tfyre.vaadin.gridexport;

import com.vaadin.data.ValueProvider;
import java.util.Collection;
import java.util.stream.Collectors;
import com.vaadin.data.provider.Query;
import com.vaadin.server.Extension;
import com.vaadin.ui.Grid;
import com.vaadin.ui.Grid.Column;
import com.vaadin.ui.UI;
import com.vaadin.ui.renderers.Renderer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD <tfyre@tfyre.co.za>
 * @param <T> the grid bean type
 */
public class DefaultGridHolder<T> implements GridHolder<T> {

    private static final long serialVersionUID = VersionHelper.serialVersionUID;

    protected HorizontalAlignment defaultAlignment = HorizontalAlignment.LEFT;

    private boolean hierarchical = false;

    private final Grid<T> grid;
    private final Map<String, Column<T, ?>> columns = new HashMap<>();
    private final Map<String, ValueProvider<T, ?>> valueProviders = new HashMap<>();
    private final List<String> propIds = new ArrayList<>();

    public DefaultGridHolder(final Grid<T> grid) {
        this.grid = grid;
    }

    @Override
    public Grid<T> getGrid() {
        return grid;
    }

    @Override
    public void hideColumn(final String propId) {
        getPropIds().remove(propId);
    }

    @Override
    public List<String> getPropIds() {
        if (propIds.isEmpty()) {
            grid.getColumns().stream()
                    .forEach(c -> {
                        propIds.add(c.getCaption());
                        columns.put(c.getCaption(), c);
                    });
        }
        return propIds;
    }

    @Override
    public boolean isHierarchical() {
        return hierarchical;
    }

    @Override
    final public void setHierarchical(final boolean hierarchical) {
        this.hierarchical = hierarchical;
    }

    @Override
    public HorizontalAlignment getCellAlignment(final String propId) {
        final Renderer<?> renderer = getRenderer(propId);
        if (renderer != null) {
            if (ExcelExport.isNumeric(renderer.getPresentationType())) {
                return HorizontalAlignment.RIGHT;
            }
        }
        return defaultAlignment;
    }

    @Override
    public boolean isColumnCollapsed(final String propertyId) {
        return getColumn(propertyId).isHidden();
    }

    @Override
    public UI getUI() {
        return grid.getUI();
    }

    @Override
    public String getColumnHeader(final String propertyId) {
        return getColumn(propertyId).getCaption();
    }

    protected Column<T, ?> getColumn(final String propId) {
        return columns.get(propId);
    }

    protected Renderer<?> getRenderer(final String propId) {
        // Grid.Column (as of 8.0.3) does not expose its renderer, we have to get it from extensions
        final Column<T, ?> column = getColumn(propId);
        if (column != null) {
            for (Extension each : column.getExtensions()) {
                if (each instanceof Renderer<?>) {
                    return (Renderer<?>) each;
                }
            }
        }
        return null;
    }

    @Override
    public Class<?> getPropertyType(final String propId) {
        Renderer<?> renderer = getRenderer(propId);
        if (renderer != null) {
            return renderer.getPresentationType();
        } else {
            return String.class;
        }
    }

    @Override
    public void setColumnValueProvider(final String propId, final ValueProvider<T, ?> valueProvider) {
        valueProviders.put(propId, valueProvider);
    }

    @Override
    public Object getPropertyValue(final T itemId, final String propId) {
        if (valueProviders.containsKey(propId)) {
            return valueProviders.get(propId).apply(itemId);
        }
        return getColumn(propId).getValueProvider().apply(itemId);
    }

    @Override
    public Collection<T> getItemIds() {
        return grid.getDataProvider().fetch(new Query<>()).collect(Collectors.toList());
    }

    @Override
    public Collection<T> getRootItemIds() {
        return getItemIds();
    }

}
