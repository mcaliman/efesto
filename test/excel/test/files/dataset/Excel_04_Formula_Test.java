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

package excel.test.files.dataset;

import excel.ExcelToolkitCommand;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * @author Massimo Caliman
 */
class Excel_04_Formula_Test {
    /**
     * ToFormula: OK
     * ToFunctional: OK
     *
     * @throws Exception
     */
    @Test
    void testTest() throws Exception {
        ExcelToolkitCommand toolkitCommand = new ExcelToolkitCommand("test/04-formula.xlsx");
        toolkitCommand.execute();
        /*toolkitCommand.writer("test/04-formula.vb");
        System.out.println("ToFormula.");
        System.out.println("-------------");
        toolkitCommand.print();
        assertTrue(toolkitCommand.test(0, "FORMULA!A1 = 1+2"));*/
        System.out.println("ToFormula.");
        System.out.println("-------------");
        toolkitCommand.toFunctional();
        assertTrue(toolkitCommand.testToFunctional(
                0, "A1 = 1+2"
        ));
        toolkitCommand.writerFormula("test/04-formula.vb");
    }
}
