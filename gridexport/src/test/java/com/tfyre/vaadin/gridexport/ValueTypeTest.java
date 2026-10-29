package com.tfyre.vaadin.gridexport;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Tests for the numeric flag of {@link ValueType}.
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD {@literal <tfyre@tfyre.co.za>}
 */
class ValueTypeTest {

    @Test
    void integerBaseIsNumeric() {
        assertTrue(ValueType.INTEGERBASE.isNumeric());
    }

    @Test
    void floatBaseIsNumeric() {
        assertTrue(ValueType.FLOATBASE.isNumeric());
    }

    @Test
    void dateTimeIsNotNumeric() {
        assertFalse(ValueType.DATETIME.isNumeric());
    }

    @Test
    void stringIsNotNumeric() {
        assertFalse(ValueType.STRING.isNumeric());
    }

}
