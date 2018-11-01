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

import efesto.parsers.StartList;
import excel.ExcelToolkitCommand;
import excel.ToolkitOptions;
import excel.grammar.Start;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * @author Massimo Caliman
 */
public class Excel_Comments_Test {

    @Test
    public void testTest() throws Exception {
        ToolkitOptions options = new ToolkitOptions();
        options.setMetadata(true);
        options.setVerbose(false);
        ExcelToolkitCommand toolkitCommand = new ExcelToolkitCommand("test/comments.xlsx", options);
        toolkitCommand.execute();
        toolkitCommand.print();
        toolkitCommand.writer("test/comments.vb");
        StartList ordered = toolkitCommand.getStartList();
        for (Start start : ordered) System.out.println(start.toString(true) + " '' " + start.getComment());
        boolean correct = toolkitCommand.test(1,
                "Sheet1!A1 = 1.0",
                "Sheet1!A2 = Sheet1!A1+1");
        assertTrue(correct);
    }
}
