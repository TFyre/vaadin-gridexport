/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.tfyre.vaadin.gridexport;

/**
 *
 * @author fsteyn
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
