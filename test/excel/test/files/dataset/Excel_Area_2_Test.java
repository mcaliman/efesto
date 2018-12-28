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
class Excel_Area_2_Test {
    /**
     * ToFormula: OK
     * ToFunctional: KO
     *
     * @throws Exception
     */
    @Test
    void testTest() throws Exception {
        ExcelToolkitCommand toolkitCommand = new ExcelToolkitCommand("test/area-2.xlsx");
        toolkitCommand.execute();
        /*toolkitCommand.writer("test/area-2.vb");
        toolkitCommand.print();
        boolean correct = toolkitCommand.test(0,
                "Area1!A1:B3 = [[11.0 21.0][12.0 22.0][13.0 23.0]]",
                "Area2!Area2Name = [[11.0 21.0][12.0 22.0][13.0 23.0][14.0 24.0]]",
                "UseArea1AndArea2!A2 = INDEX(Area1!A1:B3,2,2)",
                "UseArea1AndArea2!A1 = INDEX(Area2!Area2Name,1,2)");
        assertTrue(correct);*/
        System.out.println("ToFormula.");
        System.out.println("-------------");
        toolkitCommand.toFunctional();
        assertTrue(toolkitCommand.testToFunctional(
                0,
                "Area1!A1:B3 = [[11.0 21.0][12.0 22.0][13.0 23.0]]",
                "Area2!Area2Name = [[11.0 21.0][12.0 22.0][13.0 23.0][14.0 24.0]]",
                "UseArea1AndArea2!A2 = INDEX(Area1!A1:B3,2,2)",
                "UseArea1AndArea2!A1 = INDEX(Area2!Area2Name,1,2)"
        ));
        toolkitCommand.writerFormula("test/area-2.vb");
    }
}
