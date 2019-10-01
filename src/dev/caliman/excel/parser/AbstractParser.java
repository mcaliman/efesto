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

package dev.caliman.excel.parser;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA;

public abstract class AbstractParser {

    protected Workbook xlsxBook;
    protected Sheet xlsxSheet;//(Work)Sheet
    protected XSSFEvaluationWorkbook evalBook;

    protected boolean singleSheet;//is single xlsxSheet or not?

    protected void analyze() {
        System.out.println("Analyze...");
        this.evalBook=XSSFEvaluationWorkbook.create((XSSFWorkbook) this.xlsxBook);
        this.singleSheet=this.xlsxBook.getNumberOfSheets()==1;
    }

    protected Ptg[] tokens(Sheet sheet, int rowFormula, int colFormula) {
        int sheetIndex=this.xlsxBook.getSheetIndex(sheet);
        var sheetName=sheet.getSheetName();
        var evalSheet=evalBook.getSheet(sheetIndex);
        Ptg[] ptgs=null;
        try {
            ptgs=evalBook.getFormulaTokens(evalSheet.getCell(rowFormula, colFormula));
        } catch(FormulaParseException e) {
            System.err.println(""+e.getMessage()+sheetName+rowFormula+colFormula);
        }
        return ptgs;
    }

    protected int getSheetIndex() {
        return this.xlsxBook.getSheetIndex(this.xlsxSheet);
    }

    protected String getSheetName() {
        return this.xlsxSheet.getSheetName();
    }


    protected String getSheetName(Cell xlsxCell) {
        return xlsxCell.getSheet().getSheetName();
    }

    protected boolean isFormula(Cell xlsxCell) {
        return xlsxCell.getCellType()==CELL_TYPE_FORMULA;
    }

    protected boolean empty(final Cell xlsxCell) {
        if(xlsxCell==null) { // use row.getCell(x, Row.CREATE_NULL_AS_BLANK) to avoid null cells
            return true;
        }
        if(xlsxCell.getCellType()==Cell.CELL_TYPE_BLANK) {
            return true;
        }
        return xlsxCell.getCellType()==Cell.CELL_TYPE_STRING&&xlsxCell.getStringCellValue().trim().isEmpty();
    }

}
