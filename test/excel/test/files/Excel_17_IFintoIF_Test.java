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

class Excel_17_IFintoIF_Test {
    /**
     * ToFormula: OK
     * ToFunctional: OK
     *
     * @throws Exception
     */
    @Test
    void testTest() throws Exception {
        ExcelToolkitCommand toolkitCommand = new ExcelToolkitCommand("test/17-IF-into-IF.xlsx");
        toolkitCommand.execute();
        toolkitCommand.writer("test/17-IF-into-IF.vb");
        System.out.println("ToFormula.");
        System.out.println("-------------");
        toolkitCommand.print();
        boolean correct = toolkitCommand.test(0, "IFIF!B1 = 2.0",
                "IFIF!C1 = TRUE",
                "IFIF!A1 = IF(IFIF!B1>1,0,IF(IFIF!C1,\"Hello\",\"Bye\"))");
        assertTrue(correct);
        System.out.println("ToFunctional.");
        System.out.println("-------------");
        toolkitCommand.toFunctional();
    }
}
