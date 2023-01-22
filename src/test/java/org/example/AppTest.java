package org.example;

import junit.framework.TestCase;


import java.util.List;

import static org.example.App.findNumbersFromExcel;

/**
 * Unit test for simple App.
 */
public class AppTest
        extends TestCase {

    /**
     * Тестирование метода FindNumbersFromExcel
     */
    public void testFindNumbersFromExcel() {
        List<String> actualList = List.of("Номер QW56655 не найден", "Номер С405ММ799 найден", "Номер C052AM799 найден");
        assertEquals(findNumbersFromExcel("name_java.xlsx", "QW56655", "С405ММ799", "C052AM799"), actualList);
    }
}
