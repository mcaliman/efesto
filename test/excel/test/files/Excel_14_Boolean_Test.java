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

package excel.test.files;

import excel.ExcelToolkitCommand;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertTrue;

public class Excel_14_Boolean_Test {

    @Test
    public void testTest() throws Exception {
        ExcelToolkitCommand toolkitCommand = new ExcelToolkitCommand("test/14-Boolean.xlsx");
        toolkitCommand.execute();
        toolkitCommand.writer("test/14-Boolean.vb");
        toolkitCommand.print();
        assertTrue(toolkitCommand.test(0,
                "Boolean!A3 = 1.0",
                "Boolean!A4 = TRUE",
                "Boolean!A5 = \"IFTRUE\"",
                "Boolean!A6 = \"IFFALSE\"",
                "Boolean!A1 = IF(AND(Boolean!A3=1,Boolean!A4=TRUE),Boolean!A5,Boolean!A6)"));
    }
}
