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

class Excel_902_TerminalsFormulas_Test {
    /**
     * ToFormula: OK
     * ToFunctional: OK
     *
     * @throws Exception
     */
    @Test
    void testTest() throws Exception {
        long t = System.currentTimeMillis();
        ExcelToolkitCommand toolkitCommand = new ExcelToolkitCommand("test/902-terminals-formulas.xlsx");
        toolkitCommand.execute();
        /*toolkitCommand.writer("test/902-terminals-formulas.vb");
        toolkitCommand.print();
        boolean correct = toolkitCommand.test(0,
                "Foglio1!A1 = TRUE",
                "Foglio1!A5 = \"1/0\"",
                "Foglio1!A3 = 24.0",
                "Foglio1!A4 = \"This is a string\"",
                "Foglio1!A2 = 1.838226",
                "Foglio1!A9 = IF(Foglio1!A1,Foglio1!A5,Foglio1!A3)",
                "Foglio1!A7 = IF(Foglio1!A1,Foglio1!A2,Foglio1!A3)",
                "Foglio1!A8 = IF(Foglio1!A1,Foglio1!A4,Foglio1!A7)");
        assertTrue(correct);*/
        long elapsed = System.currentTimeMillis() - t;
        System.out.println("elapsed: " + elapsed / 1000 + " s.");
        System.out.println("ToFormula.");
        toolkitCommand.toFunctional();
        assertTrue(toolkitCommand.testToFunctional(
                0,
                "A1 = TRUE",
                "A5 = \"1/0\"",
                "A3 = 24.0",
                "A4 = \"This is a string\"",
                "A2 = 1.838226",
                "A9 = IF(A1,A5,A3)",
                "A7 = IF(A1,A2,A3)",
                "A8 = IF(A1,A4,A7)"
        ));
        toolkitCommand.writerFormula("test/902-terminals-formulas.vb");
    }

}
