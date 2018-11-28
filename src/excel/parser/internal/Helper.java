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

package excel.parser.internal;

import excel.grammar.Start;
import org.apache.poi.ss.formula.EvaluationSheet;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Helper {

    Workbook workbook;
    XSSFEvaluationWorkbook evalBook;

    public Helper(Workbook workbook) {
        this.workbook = workbook;
        this.evalBook = XSSFEvaluationWorkbook.create((XSSFWorkbook) workbook);
    }

    public int getSheetIndex(String sheetName){
        int sheetIndex = evalBook.getSheetIndex(sheetName);
        return sheetIndex;
    }

    public  Ptg[] tokens(Sheet sheet, int rowFormula, int colFormula) {
        int sheetIndex = workbook.getSheetIndex(sheet);
        var sheetName = sheet.getSheetName();
        var evalSheet = evalBook.getSheet(sheetIndex);
        Ptg[] ptgs = null;
        try {
            ptgs = evalBook.getFormulaTokens(evalSheet.getCell(rowFormula, colFormula));
        } catch (FormulaParseException e) {
            err("" + e.getMessage(),sheetName, rowFormula, colFormula);
        }
        return ptgs;
    }

    private void err(String string, String sheetName,int row, int column) {
        System.err.println(Start.cellAddress(row, column, sheetName) + " parse error: " + string);
    }
}
