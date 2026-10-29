package com.tfyre.vaadin.gridexport;

import com.vaadin.flow.component.grid.Grid;
import java.util.Date;
import java.util.List;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Tests for {@link DefaultGridHolder}.
 *
 * @author Francois Steyn - TFyreIT (PTY) LTD {@literal <tfyre@tfyre.co.za>}
 */
class DefaultGridHolderTest {

    private static final Date HIRE_DATE = new Date(1_700_000_000_000L);
    private static final Person ALICE = new Person("Alice", 30, 1234.5);
    private static final Person BOB = new Person("Bob", 41, 987.65);

    private Grid<Person> grid;
    private DefaultGridHolder<Person> holder;

    @BeforeEach
    void setUp() {
        grid = new Grid<>();
        grid.setItems(List.of(ALICE, BOB));
        holder = new DefaultGridHolder<>(grid)
                .addColumn("name", grid.addColumn(Person::name), Person::name, "Name", ValueType.STRING)
                .addColumn("age", grid.addColumn(Person::age), Person::age, "Age", ValueType.INTEGERBASE)
                .addColumn("salary", grid.addColumn(Person::salary), Person::salary, "Salary", ValueType.FLOATBASE)
                .addColumn("hired", grid.addColumn(person -> HIRE_DATE), person -> HIRE_DATE, "Hired", ValueType.DATETIME);
    }

    @Test
    void getGridReturnsWrappedGrid() {
        assertSame(grid, holder.getGrid());
    }

    @Test
    void addColumnKeepsRegistrationOrder() {
        assertEquals(List.of("name", "age", "salary", "hired"), holder.getColumnKeys());
    }

    @Test
    void getColumnKeysReturnsUnmodifiableList() {
        final List<String> keys = holder.getColumnKeys();
        assertThrows(UnsupportedOperationException.class, () -> keys.add("other"));
    }

    @Test
    void getColumnHeaderReturnsRegisteredHeaderName() {
        assertEquals("Name", holder.getColumnHeader("name"));
        assertEquals("Age", holder.getColumnHeader("age"));
        assertEquals("Salary", holder.getColumnHeader("salary"));
        assertEquals("Hired", holder.getColumnHeader("hired"));
    }

    @Test
    void getPropertyTypeReturnsRegisteredValueType() {
        assertEquals(ValueType.STRING, holder.getPropertyType("name"));
        assertEquals(ValueType.INTEGERBASE, holder.getPropertyType("age"));
        assertEquals(ValueType.FLOATBASE, holder.getPropertyType("salary"));
        assertEquals(ValueType.DATETIME, holder.getPropertyType("hired"));
    }

    @Test
    void getPropertyValueAppliesRegisteredValueProvider() {
        assertEquals("Alice", holder.getPropertyValue(ALICE, "name"));
        assertEquals(30, holder.getPropertyValue(ALICE, "age"));
        assertEquals(1234.5, holder.getPropertyValue(ALICE, "salary"));
        assertSame(HIRE_DATE, holder.getPropertyValue(ALICE, "hired"));

        assertEquals("Bob", holder.getPropertyValue(BOB, "name"));
        assertEquals(41, holder.getPropertyValue(BOB, "age"));
        assertEquals(987.65, holder.getPropertyValue(BOB, "salary"));
    }

    @Test
    void numericColumnsAreRightAligned() {
        assertEquals(HorizontalAlignment.RIGHT, holder.getCellAlignment("age"));
        assertEquals(HorizontalAlignment.RIGHT, holder.getCellAlignment("salary"));
    }

    @Test
    void nonNumericColumnsAreLeftAligned() {
        assertEquals(HorizontalAlignment.LEFT, holder.getCellAlignment("name"));
        assertEquals(HorizontalAlignment.LEFT, holder.getCellAlignment("hired"));
    }

    @Test
    void getItemsReturnsAllItemsFromDataProvider() {
        assertEquals(List.of(ALICE, BOB), holder.getItems());
    }

    @Test
    void clearColumnsRemovesAllRegisteredColumns() {
        holder.clearColumns();
        assertTrue(holder.getColumnKeys().isEmpty());
    }

    private record Person(String name, int age, double salary) {
    }

}
