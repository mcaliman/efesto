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
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.EvaluationName;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.ptg.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.function.Predicate;

import static java.lang.System.err;
import static org.apache.poi.ss.usermodel.Cell.*;

public abstract class AbstractParser {

    protected final Predicate<Ptg> arrayPtg = (Ptg t) -> t instanceof ArrayPtg;
    protected final Predicate<Ptg> addPtg = (Ptg t) -> t instanceof AddPtg;
    protected final Predicate<Ptg> area3DPxg = (Ptg t) -> t instanceof Area3DPxg;
    protected final Predicate<Ptg> areaErrPtg = (Ptg t) -> t instanceof AreaErrPtg;
    protected final Predicate<Ptg> areaPtg = (Ptg t) -> t instanceof AreaPtg;
    protected final Predicate<Ptg> attrPtg = (Ptg t) -> t instanceof AttrPtg;
    protected final Predicate<Ptg> boolPtg = (Ptg t) -> t instanceof BoolPtg;
    protected final Predicate<Ptg> concatPtg = (Ptg t) -> t instanceof ConcatPtg;
    protected final Predicate<Ptg> deleted3DPxg = (Ptg t) -> t instanceof Deleted3DPxg;
    protected final Predicate<Ptg> deletedArea3DPtg = (Ptg t) -> t instanceof DeletedArea3DPtg;
    protected final Predicate<Ptg> deletedRef3DPtg = (Ptg t) -> t instanceof DeletedRef3DPtg;
    protected final Predicate<Ptg> dividePtg = (Ptg t) -> t instanceof DividePtg;
    protected final Predicate<Ptg> equalPtg = (Ptg t) -> t instanceof EqualPtg;
    protected final Predicate<Ptg> errPtg = (Ptg t) -> t instanceof ErrPtg;
    protected final Predicate<Ptg> funcPtg = (Ptg t) -> t instanceof FuncPtg;
    protected final Predicate<Ptg> funcVarPtg = (Ptg t) -> t instanceof FuncVarPtg;
    protected final Predicate<Ptg> greaterEqualPtg = (Ptg t) -> t instanceof GreaterEqualPtg;
    protected final Predicate<Ptg> greaterThanPtg = (Ptg t) -> t instanceof GreaterThanPtg;
    protected final Predicate<Ptg> intersectionPtg = (Ptg t) -> t instanceof IntersectionPtg;
    protected final Predicate<Ptg> intPtg = (Ptg t) -> t instanceof IntPtg;
    protected final Predicate<Ptg> lessEqualPtg = (Ptg t) -> t instanceof LessEqualPtg;
    protected final Predicate<Ptg> lessThanPtg = (Ptg t) -> t instanceof LessThanPtg;
    protected final Predicate<Ptg> memErrPtg = (Ptg t) -> t instanceof MemErrPtg;
    protected final Predicate<Ptg> missingArgPtg = (Ptg t) -> t instanceof MissingArgPtg;
    protected final Predicate<Ptg> multiplyPtg = (Ptg t) -> t instanceof MultiplyPtg;
    protected final Predicate<Ptg> namePtg = (Ptg t) -> t instanceof NamePtg;
    protected final Predicate<Ptg> notEqualPtg = (Ptg t) -> t instanceof NotEqualPtg;
    protected final Predicate<Ptg> numberPtg = (Ptg t) -> t instanceof NumberPtg;
    protected final Predicate<Ptg> parenthesisPtg = (Ptg t) -> t instanceof ParenthesisPtg;
    protected final Predicate<Ptg> percentPtg = (Ptg t) -> t instanceof PercentPtg;
    protected final Predicate<Ptg> powerPtg = (Ptg t) -> t instanceof PowerPtg;
    protected final Predicate<Ptg> ref3DPxg = (Ptg t) -> t instanceof Ref3DPxg;
    protected final Predicate<Ptg> refErrorPtg = (Ptg t) -> t instanceof RefErrorPtg;
    protected final Predicate<Ptg> refPtg = (Ptg t) -> t instanceof RefPtg;
    protected final Predicate<Ptg> stringPtg = (Ptg t) -> t instanceof StringPtg;
    protected final Predicate<Ptg> subtractPtg = (Ptg t) -> t instanceof SubtractPtg;
    protected final Predicate<Ptg> unaryMinusPtg = (Ptg t) -> t instanceof UnaryMinusPtg;
    protected final Predicate<Ptg> unaryPlusPtg = (Ptg t) -> t instanceof UnaryPlusPtg;
    protected final Predicate<Ptg> unionPtg = (Ptg t) -> t instanceof UnionPtg;
    protected final Predicate<Ptg> unknownPtg = (Ptg t) -> t instanceof UnknownPtg;

    protected String filename;
    protected File file;

    protected Workbook workbook;
    protected Sheet sheet;
    protected XSSFEvaluationWorkbook evaluation;

    protected Ptg[] formulaPtgs;
    protected String formulaAddress;
    protected String formulaPlainText;
    protected int noOfFormulas;//formula counters noOfFormulas
    protected int noOfSheets;

    protected boolean singleSheet;//is single sheet or not?

    protected int column;//Current Formula Column
    protected int row;//Current Formula Row

    protected AbstractParser(String filename) throws IOException, InvalidFormatException {
        this.filename = filename;
        this.file = new File(this.filename);
        this.workbook = WorkbookFactory.create(this.file);
    }

    public String getFilename() {
        return this.filename;
    }

    public void parse() {
        this.evaluation = XSSFEvaluationWorkbook.create((XSSFWorkbook) this.workbook);
        this.noOfSheets = this.workbook.getNumberOfSheets();
        this.singleSheet = this.noOfSheets == 1;
        for(Sheet sheet : this.workbook) {
            this.sheet = sheet;
            for(Row row : this.sheet)
                for(Cell cell : row) if(!empty(cell)) parse(cell);
        }
    }


    protected abstract void parse(Cell cell);


    protected void parseFormula(Cell cell) {
        this.noOfFormulas++;
        this.column = cell.getColumnIndex();
        this.row = cell.getRowIndex();
        this.formulaAddress = cellAddress();
        this.formulaPlainText = cell.getCellFormula();
        System.out.println("Formula Plain Text: " + this.formulaAddress);
        this.formulaPtgs = tokens();

    }

    protected Ptg[] tokens() {
        int sheetIndex = this.getSheetIndex();
        var sheetName = this.getSheetName();
        var evaluationSheet = this.evaluation.getSheet(sheetIndex);
        Ptg[] ptgs = null;
        try {
            EvaluationCell evaluationCell = evaluationSheet.getCell(this.row, this.column);
            ptgs = this.evaluation.getFormulaTokens(evaluationCell);
        } catch(FormulaParseException e) {
            err.println("" + e.getMessage() + sheetName + this.row + this.column);
        }
        return ptgs;
    }

    protected Ptg[] getName(NamePtg t) {
        EvaluationName evaluationName = this.evaluation.getName(t);
        return evaluationName.getNameDefinition();
    }

    protected String getNameText(NamePtg t) {
        return this.evaluation.getNameText(t);
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

    protected String getSheetName(Cell cell) {
        return cell.getSheet().getSheetName();
    }

    protected boolean isFormula(final Cell cell) {
        return cell.getCellType() == CELL_TYPE_FORMULA;
    }

    protected boolean empty(final Cell cell) {
        if(cell == null) return true;
        if(cell.getCellType() == Cell.CELL_TYPE_BLANK) return true;
        return cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().trim().isEmpty();
    }

    protected String getCellAddress() {
        return Start.cellAddress(this.row, this.column, this.getSheetName());
    }

    protected void doesFormulaReferToDeletedCell() {
        err.println(getCellAddress() + " does formula refer to deleted cell");
    }

    protected void parseErrPtg(Ptg t) {
        err.println(t.getClass().getName() + ": " + t.toString());
    }

    protected void parseMissingArguments() {
        err.println("Missing ExcelFunction Arguments for cell: " + getCellAddress());
    }

    protected Object parseCellValue(Cell cell) {
        if(cell == null) return null;
        if(isDataType(cell))
            return cell.getDateCellValue();
        switch(cell.getCellType()) {
            case CELL_TYPE_STRING:
            case CELL_TYPE_BLANK:
                return cell.getStringCellValue();
            case CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case CELL_TYPE_FORMULA:
                if(cell.toString() != null && cell.toString().equalsIgnoreCase("TRUE")) {
                    return true;
                }
                if(cell.toString() != null && cell.toString().equalsIgnoreCase("FALSE")) {
                    return false;
                }
                return cell.toString();
            default:
                return null;
        }
    }

    private boolean isDataType(Cell cell) {
        return cell.getCellType() == CELL_TYPE_NUMERIC && HSSFDateUtil.isCellDateFormatted(cell);
    }
}
