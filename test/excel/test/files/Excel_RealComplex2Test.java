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
class Excel_RealComplex2Test {

    @Test
    void testTest() throws Exception {
        if (true) {
            long t = System.currentTimeMillis();
            ToolkitOptions options = new ToolkitOptions();
            options.setMetadata(false);
            ExcelToolkitCommand toolkitCommand = new ExcelToolkitCommand("D:/xl2.xlsx", options);
            System.out.println("execute()...");
            toolkitCommand.execute();
            System.out.println("writer()...");
            toolkitCommand.writer("D:/xl2.vb");
            long elapsed = System.currentTimeMillis() - t;
            System.out.println("Elapsed time: " + elapsed + " [ms] or " + elapsed / 1000 + " [s]. or " + elapsed / 1000 / 60 + "[m].");
            //Elapsed time: 522682 [ms] or 522 [s]. or 8[m]. release 201807231738
            //Elapsed time: 553386 [ms] or 553 [s]. or 9[m]. release 201807241325
            //Elapsed time: 587383 [ms] or 587 [s]. or 9[m]. release 201807261612
            //Elapsed time: 971285 [ms] or 971 [s]. or 16[m].


            //System.out.println("Call transpiler.");
            //toolkitCommand.transpiler();
        }

    }

}
