/*
 * Efesto - Excel Formula Extractor System and Topological Ordering algorithm.
 * Copyright (C) 2017 Massimo Caliman mcaliman@caliman.biz
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as published
 * by the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 *
 * If AGPL Version 3.0 terms are incompatible with your use of
 * Efesto, alternative license terms are available from Massimo Caliman
 * please direct inquiries about Efesto licensing to mcaliman@caliman.biz
 */

package excel.test;

import excel.parser.internal.HelperInternal;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * @author mcaliman
 */
class ExcelHelperTest {

    ExcelHelperTest() {
    }

    /**
     * Test of columnAsLetter method, of class ExcelHelper.
     */
    @Test
    void testColumnAsLetter() {
        System.out.println("columnAsLetter");

        assertEquals("A", HelperInternal.columnAsLetter(0));
        assertEquals("B", HelperInternal.columnAsLetter(1));
    }

    /**
     * Test of cellAddress method, of class ExcelHelper.
     */
    @Test
    void testCellAddress_int_int() {
        System.out.println("cellAddress(row,col)");

        assertEquals("A1", HelperInternal.cellAddress(0, 0));

    }

    /**
     * Test of cellAddress method, of class ExcelHelper.
     */
    @Test
    void testCellAddress_3args() {
        System.out.println("cellAddress(row,col,sheetname)");

        String result = HelperInternal.cellAddress(0, 0, "Sheet");
        System.out.println(result);
        assertEquals("Sheet!A1", result);
    }

}
