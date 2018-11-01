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

package efesto.internals;

import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.EvaluationSheet;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;

public class FormulaTokensInternal {


    private final XSSFEvaluationWorkbook ew;
    private final EvaluationSheet es;

    public FormulaTokensInternal(XSSFEvaluationWorkbook ew, EvaluationSheet es) {
        this.ew = ew;
        this.es = es;
    }

    public Ptg[] getFormulaTokens(int row, int column) {
        EvaluationCell evalCell = es.getCell(row, column);
        Ptg[] ptgs = null;
        try {
            ptgs = ew.getFormulaTokens(evalCell);
        } catch (FormulaParseException e) {
            err("" + e.getMessage(), row, column);
        }
        return ptgs;
    }

    private void err(String string, int row, int column) {
        System.err.println(string + " row:" + row + " col:" + column);
    }
}
