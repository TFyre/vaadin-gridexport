package com.tfyre.vaadin.gridexport;

import java.io.InputStream;
import java.io.Serializable;
import java.util.List;

/**
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD {@literal <tfyre@tfyre.co.za>}
 * @param <T> the grid bean type
 */
public abstract class GridExport<T> implements Serializable {

    private static final long serialVersionUID = VersionHelper.serialVersionUID;

    private GridHolder<T> gridHolder;

    public GridExport(final GridHolder<T> gridHolder) {
        this.gridHolder = gridHolder;
    }

    public GridHolder<T> getGridHolder() {
        return gridHolder;
    }

    public List<String> getPropIds() {
        return gridHolder.getPropIds();
    }

    public void setGridHolder(final GridHolder<T> gridHolder) {
        this.gridHolder = gridHolder;
    }

    public boolean isHierarchical() {
        return gridHolder.isHierarchical();
    }

    public abstract void convertGrid();

    public abstract InputStream getInputStream();

}
