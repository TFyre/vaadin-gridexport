package com.tfyre.vaadin.gridexport;

/**
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD {@literal <tfyre@tfyre.co.za>}
 */
public enum ValueType {
    DATETIME(false),
    INTEGERBASE(true),
    FLOATBASE(true),
    STRING(false);

    private final boolean numeric;

    private ValueType(boolean numeric) {
        this.numeric = numeric;
    }

    public boolean isNumeric() {
        return numeric;
    }

}
