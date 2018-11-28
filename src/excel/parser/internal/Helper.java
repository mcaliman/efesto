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
import org.apache.poi.ss.formula.EvaluationName;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.ptg.Area3DPxg;
import org.apache.poi.ss.formula.ptg.NamePtg;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.Ref3DPxg;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class Helper {

    private final Workbook workbook;
    final XSSFEvaluationWorkbook evalBook;

    public Helper(Workbook workbook) {
        this.workbook = workbook;
        this.evalBook = XSSFEvaluationWorkbook.create((XSSFWorkbook) workbook);
    }

    public Ptg[] getName(NamePtg t){
        EvaluationName evaluationName = evalBook.getName(t);
        return evaluationName.getNameDefinition();
    }

    public String getNameText(NamePtg t){
        return evalBook.getNameText(t);
    }

    public String getArea(Area3DPxg t){
        return t.format2DRefAsString();
    }

    public String getCellRef(Ref3DPxg t){
        return t.format2DRefAsString();
    }

    public int getSheetIndex(String sheetName){
        return evalBook.getSheetIndex(sheetName);
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
