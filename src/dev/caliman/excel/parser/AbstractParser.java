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

import dev.caliman.excel.grammar.Start;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.ptg.ArrayPtg;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.function.Predicate;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA;

public abstract class AbstractParser {

    protected final Predicate<Ptg> arrayPtg = (Ptg t) -> t instanceof ArrayPtg;

    protected String fileName;
    protected File file;

    protected Workbook workbook;
    protected Sheet sheet;
    protected XSSFEvaluationWorkbook evaluation;

    protected Ptg[] formulaPtgs;
    protected String formulaAddress;
    protected String formulaPlainText;


    protected boolean singleSheet;//is single sheet or not?
    protected int counterFormulas;//formula counters
    protected int column;//Current Formula Column
    protected int row;//Current Formula Row

    protected AbstractParser(String xlsxFileName) throws IOException, InvalidFormatException {
        this.fileName = xlsxFileName;
        this.file = new File(this.fileName);
        this.workbook = WorkbookFactory.create(file);
    }

    public String getFileName() {
        return fileName;
    }

    public void parse() {
        analyze();
        for(Sheet currentSheet : this.workbook) {
            this.sheet = currentSheet;
            parseSheet();
        }
    }

    protected void analyze() {
        System.out.println("Analyze...");
        this.evaluation = XSSFEvaluationWorkbook.create((XSSFWorkbook) this.workbook);
        this.singleSheet = this.workbook.getNumberOfSheets() == 1;
    }

    protected void parseSheet() {
        for(Row xlsxRow : sheet)
            for(Cell xlsxCell : xlsxRow)
                if(!empty(xlsxCell)) parse(xlsxCell);
                else {
                    System.err.println("Cell is null.");
                    //throw new RuntimeException("Cell is null.");
                }
    }


    protected abstract void parse(Cell xlsxCell);


    protected void parseFormula(Cell xlsxCell) {
        this.counterFormulas++;
        this.column = xlsxCell.getColumnIndex();
        this.row = xlsxCell.getRowIndex();
        this.formulaAddress = cellAddress();
        this.formulaPlainText = xlsxCell.getCellFormula();
        System.out.println("Formula Plain Text: " + this.formulaAddress);
        this.formulaPtgs = tokens(this.sheet, this.row, this.column);

    }

    protected Ptg[] tokens(Sheet sheet, int rowFormula, int colFormula) {
        int sheetIndex = this.workbook.getSheetIndex(sheet);
        var sheetName = sheet.getSheetName();
        var evalSheet = evaluation.getSheet(sheetIndex);
        Ptg[] ptgs = null;
        try {
            ptgs = evaluation.getFormulaTokens(evalSheet.getCell(rowFormula, colFormula));
        } catch(FormulaParseException e) {
            System.err.println("" + e.getMessage() + sheetName + rowFormula + colFormula);
        }
        return ptgs;
    }

    protected String cellAddress() {
        return Start.cellAddress(this.row, this.column, getSheetName());
    }

    protected int getSheetIndex() {
        return this.workbook.getSheetIndex(this.sheet);
    }

    protected String getSheetName() {
        return this.sheet.getSheetName();
    }

    protected String getSheetName(Cell xlsxCell) {
        return xlsxCell.getSheet().getSheetName();
    }

    protected boolean isFormula(Cell xlsxCell) {
        return xlsxCell.getCellType() == CELL_TYPE_FORMULA;
    }

    protected boolean empty(final Cell xlsxCell) {
        if(xlsxCell == null) return true;
        if(xlsxCell.getCellType() == Cell.CELL_TYPE_BLANK) return true;
        return xlsxCell.getCellType() == Cell.CELL_TYPE_STRING && xlsxCell.getStringCellValue().trim().isEmpty();
    }

}
