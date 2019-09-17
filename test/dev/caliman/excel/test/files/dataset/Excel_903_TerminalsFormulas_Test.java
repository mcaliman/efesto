/*
 * Efesto - Excel Formula Extractor System and Topological Ordering algorithm.
 * Copyright (C) 2017 Massimo Caliman mcaliman@gmail.com
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
 * please direct inquiries about Efesto licensing to mcaliman@gmail.com
 */

package dev.caliman.excel.test.files.dataset;

import dev.caliman.excel.ToolkitCommand;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertTrue;

class Excel_903_TerminalsFormulas_Test {

    @Test
    void testTest() throws Exception {
        ToolkitCommand toolkitCommand = new ToolkitCommand("test/903-terminals-formulas.xlsx");
        toolkitCommand.execute();
        System.out.println("ToFormula.");
        toolkitCommand.toFormula();
        assertTrue(toolkitCommand.testToFormula(
                0,
                "A1 = TRUE",
                "A2 = 1.838226",
                "A3 = 24.0",
                "A4 = \"This is a string\"",
                "A6 = IF(A1,A2,A3)",
                "A7 = IF(A1,A4,A3)"
        ));
        toolkitCommand.write("test/903-terminals-formulas.vb");
    }

}
