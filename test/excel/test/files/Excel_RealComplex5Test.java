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
import excel.ToolkitOptions;
import org.junit.jupiter.api.Test;

/**
 * @author mcaliman
 */
@SuppressWarnings("UnusedAssignment")
class Excel_RealComplex5Test {

    @Test
    void testTest() throws Exception {
        if (true) {
            long t = System.currentTimeMillis();
            ToolkitOptions options = new ToolkitOptions();
            options.setMetadata(false);
            options.setVerbose(false);
            ExcelToolkitCommand toolkitCommand = new ExcelToolkitCommand("D:/xl5.xlsx", options);
            System.out.println("execute()...");
            toolkitCommand.execute();
            System.out.println("writer()...");
            toolkitCommand.writer("D:/xl5.vb");
            long elapsed = System.currentTimeMillis() - t;
            System.out.println("Elapsed time: " + elapsed + " [ms] or " + elapsed / 1000 + " [s]. or " + elapsed / 1000 / 60 + "[m].");
            System.out.println("ToFunctional.");
            toolkitCommand.toFunctional();
        }

    }

}
