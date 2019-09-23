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

import dev.caliman.excel.grammar.Formula;
import dev.caliman.excel.grammar.Start;
import dev.caliman.excel.grammar.formula.ConstantArray;
import dev.caliman.excel.grammar.formula.ParenthesisFormula;
import dev.caliman.excel.grammar.formula.Reference;
import dev.caliman.excel.grammar.formula.constant.*;
import dev.caliman.excel.grammar.formula.functioncall.EXCEL_FUNCTION;
import dev.caliman.excel.grammar.formula.functioncall.PercentFormula;
import dev.caliman.excel.grammar.formula.functioncall.binary.*;
import dev.caliman.excel.grammar.formula.functioncall.builtin.SUM;
import dev.caliman.excel.grammar.formula.functioncall.unary.Minus;
import dev.caliman.excel.grammar.formula.functioncall.unary.Plus;
import dev.caliman.excel.grammar.formula.reference.*;
import dev.caliman.excel.grammar.formula.reference.referencefunction.OFFSET;
import dev.caliman.excel.graph.StartGraph;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.ptg.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jetbrains.annotations.NotNull;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Stack;
import java.util.function.Consumer;
import java.util.function.Predicate;
import java.util.stream.Stream;

import static java.lang.System.err;
import static java.lang.System.out;
import static org.apache.poi.ss.formula.ptg.ErrPtg.*;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA;


/**
 * @author Massimo Caliman
 */
@SuppressWarnings("JavaDoc")
public final class Parser {

    private final Predicate<Ptg> arrayPtg = (Ptg t) -> t instanceof ArrayPtg;
    private final Predicate<Ptg> addPtg = (Ptg t) -> t instanceof AddPtg;
    private final Predicate<Ptg> area3DPxg = (Ptg t) -> t instanceof Area3DPxg;
    private final Predicate<Ptg> areaErrPtg = (Ptg t) -> t instanceof AreaErrPtg;
    private final Predicate<Ptg> areaPtg = (Ptg t) -> t instanceof AreaPtg;
    private final Predicate<Ptg> attrPtg = (Ptg t) -> t instanceof AttrPtg;
    private final Predicate<Ptg> boolPtg = (Ptg t) -> t instanceof BoolPtg;
    private final Predicate<Ptg> concatPtg = (Ptg t) -> t instanceof ConcatPtg;
    private final Predicate<Ptg> deleted3DPxg = (Ptg t) -> t instanceof Deleted3DPxg;
    private final Predicate<Ptg> deletedArea3DPtg = (Ptg t) -> t instanceof DeletedArea3DPtg;
    private final Predicate<Ptg> deletedRef3DPtg = (Ptg t) -> t instanceof DeletedRef3DPtg;
    private final Predicate<Ptg> dividePtg = (Ptg t) -> t instanceof DividePtg;
    private final Predicate<Ptg> equalPtg = (Ptg t) -> t instanceof EqualPtg;
    private final Predicate<Ptg> errPtg = (Ptg t) -> t instanceof ErrPtg;
    private final Predicate<Ptg> funcPtg = (Ptg t) -> t instanceof FuncPtg;
    private final Predicate<Ptg> funcVarPtg = (Ptg t) -> t instanceof FuncVarPtg;
    private final Predicate<Ptg> greaterEqualPtg = (Ptg t) -> t instanceof GreaterEqualPtg;
    private final Predicate<Ptg> greaterThanPtg = (Ptg t) -> t instanceof GreaterThanPtg;
    private final Predicate<Ptg> intersectionPtg = (Ptg t) -> t instanceof IntersectionPtg;
    private final Predicate<Ptg> intPtg = (Ptg t) -> t instanceof IntPtg;
    private final Predicate<Ptg> lessEqualPtg = (Ptg t) -> t instanceof LessEqualPtg;
    private final Predicate<Ptg> lessThanPtg = (Ptg t) -> t instanceof LessThanPtg;
    private final Predicate<Ptg> memErrPtg = (Ptg t) -> t instanceof MemErrPtg;
    private final Predicate<Ptg> missingArgPtg = (Ptg t) -> t instanceof MissingArgPtg;
    private final Predicate<Ptg> multiplyPtg = (Ptg t) -> t instanceof MultiplyPtg;
    private final Predicate<Ptg> namePtg = (Ptg t) -> t instanceof NamePtg;
    private final Predicate<Ptg> notEqualPtg = (Ptg t) -> t instanceof NotEqualPtg;
    private final Predicate<Ptg> numberPtg = (Ptg t) -> t instanceof NumberPtg;
    private final Predicate<Ptg> parenthesisPtg = (Ptg t) -> t instanceof ParenthesisPtg;
    private final Predicate<Ptg> percentPtg = (Ptg t) -> t instanceof PercentPtg;
    private final Predicate<Ptg> powerPtg = (Ptg t) -> t instanceof PowerPtg;
    private final Predicate<Ptg> ref3DPxg = (Ptg t) -> t instanceof Ref3DPxg;
    private final Predicate<Ptg> refErrorPtg = (Ptg t) -> t instanceof RefErrorPtg;
    private final Predicate<Ptg> refPtg = (Ptg t) -> t instanceof RefPtg;
    private final Predicate<Ptg> stringPtg = (Ptg t) -> t instanceof StringPtg;
    private final Predicate<Ptg> subtractPtg = (Ptg t) -> t instanceof SubtractPtg;
    private final Predicate<Ptg> unaryMinusPtg = (Ptg t) -> t instanceof UnaryMinusPtg;
    private final Predicate<Ptg> unaryPlusPtg = (Ptg t) -> t instanceof UnaryPlusPtg;
    private final Predicate<Ptg> unionPtg = (Ptg t) -> t instanceof UnionPtg;
    private final Predicate<Ptg> unknownPtg = (Ptg t) -> t instanceof UnknownPtg;
    public boolean verbose = false;
    private int colFormula;//Current Formula Column
    private int rowFormula;//Current Formula Row
    private int sheetIndex;//Current Sheet Index
    private String sheetName;//Current Sheet Name
    private boolean isSingleSheet;
    private Workbook book;
    private Helper helper;
    private List<Cell> ext;
    private int counterFormulas;
    private Sheet sheet;//(Work)Sheet
    private boolean protectionPresent;//(Work)Book Protection Present flag
    private String fileName;

    private StartList unordered;
    private StartList ordered;
    private StartGraph graph;
    private Stack<Start> stack;

    public Parser(@NotNull String filename) throws IOException, InvalidFormatException {
        File file = new File(filename);
        this.book = WorkbookFactory.create(file);
        this.ext = new ArrayList<>();
        this.helper = new Helper(this.book);
        this.fileName = file.getName();
        this.unordered = new StartList();
        this.ordered = new StartList();
        this.graph = new StartGraph();
        this.stack = new Stack<>();
    }

    public int getCounterFormulas() {
        return counterFormulas;
    }

    public String getFileName() {
        return fileName;
    }

    private void verbose(String text) {
        if ( this.verbose ) out.println(text);
    }

    private void parse(@NotNull Sheet sheet) {
        this.sheet = sheet;
        protectionPresent = protectionPresent || ((XSSFSheet) sheet).validateSheetPassword("password");
        this.sheetIndex = book.getSheetIndex(sheet);
        this.sheetName = sheet.getSheetName();
        verbose("Parsing sheet-name:" + this.sheetName);
        for (Row row : sheet)
            for (Cell cell : row)
                if ( cell != null ) parse(cell);
                else err("Cell is null.", rowFormula, colFormula);
    }

    private void parse(Cell cell) {
        if ( cell.getCellType() == CELL_TYPE_FORMULA ) {
            parseFormula(cell);
            this.counterFormulas++;
        } else if ( this.ext.contains(cell) ) {
            verbose("Recover loosed cell!");
            Object obj = Helper.valueOf(cell);
            CELL cellRef = new CELL(cell.getRowIndex(), cell.getColumnIndex());
            cellRef.setValue(obj);
            cellRef.setSheetName(cell.getSheet().getSheetName());
            cellRef.setSheetIndex(helper.getSheetIndex(cell.getSheet().getSheetName()));
            parseCELLlinked(cellRef);
            this.ext.remove(cell);
        }
    }

    private void parseFormula(Cell cell) {
        verbose("Cell:" + cell.getClass().getSimpleName() + " " + cell.toString() + " " + cell.getCellType());
        colFormula = cell.getColumnIndex();
        rowFormula = cell.getRowIndex();
        String formulaAddress = Start.cellAddress(rowFormula, colFormula);
        String text = cell.getCellFormula();
        out.println("RAW>> " + formulaAddress + " = " + text);
        Ptg[] formulaPtgs = helper.tokens(this.sheet, this.rowFormula, this.colFormula);
        if ( formulaPtgs == null ) {
            String formulaText = cell.getCellFormula();
            err.println("ptgs empty or null for address " + formulaAddress);
            err("ptgs empty or null for address " + formulaAddress, rowFormula, colFormula);
            parseUDF(formulaText);
            return;
        }
        Start start = parse(formulaPtgs);
        if ( Objects.nonNull(start) ) {
            start.setSingleSheet(this.isSingleSheet);
            parseFormula(start);
        }
    }

    private Start parse(@NotNull Ptg[] ptgs) {
        stack.empty();
        if ( Ptg.doesFormulaReferToDeletedCell(ptgs) ) doesFormulaReferToDeletedCell(rowFormula, colFormula);
        for (Ptg ptg : ptgs) parse(ptg, rowFormula, colFormula);
        Start start = null;
        if ( !stack.empty() ) start = stack.pop();
        return start;
    }

    private void parse(@NotNull Ptg p, int row, int column) {
        verbose("parse: " + p.getClass().getSimpleName());
        try (Stream<WhatIf> stream = Stream.of(
                new WhatIf(p, arrayPtg, (Ptg t) -> parseArrayPtg((ArrayPtg) t)),
                new WhatIf(p, addPtg, (Ptg t) -> parseAdd()),
                new WhatIf(p, area3DPxg, (Ptg t) -> parseArea3DPxg((Area3DPxg) t)),
                new WhatIf(p, areaErrPtg, (Ptg t) -> parseAreaErrPtg((AreaErrPtg) t)),
                new WhatIf(p, areaPtg, (Ptg t) -> parseAreaPtg((AreaPtg) t)),
                new WhatIf(p, attrPtg, (Ptg t) -> parseAttrPtg((AttrPtg) t)),
                new WhatIf(p, boolPtg, t -> parseBooleanLiteral(((BoolPtg) t).getValue())),
                new WhatIf(p, concatPtg, t -> parseConcat()),
                new WhatIf(p, deleted3DPxg, (Ptg t) -> parseDeleted3DPxg((Deleted3DPxg) t)),
                new WhatIf(p, deletedArea3DPtg, (Ptg t) -> parseDeletedArea3DPtg((DeletedArea3DPtg) t)),
                new WhatIf(p, deletedRef3DPtg, (Ptg t) -> parseDeletedRef3DPtg((DeletedRef3DPtg) t)),
                new WhatIf(p, dividePtg, t -> parseDiv()),
                new WhatIf(p, equalPtg, t -> parseEq()),
                new WhatIf(p, errPtg, (Ptg t) -> parseErrorLiteral((ErrPtg) t)),
                new WhatIf(p, funcPtg, (Ptg t) -> parseFuncPtg((FuncPtg) t)),
                new WhatIf(p, funcVarPtg, (Ptg t) -> parseFuncVarPtg((FuncVarPtg) t)),
                new WhatIf(p, greaterEqualPtg, t -> parseGteq()),
                new WhatIf(p, greaterThanPtg, t -> parseGt()),
                new WhatIf(p, intersectionPtg, t -> parseIntersection()),
                new WhatIf(p, intPtg, t -> parseIntLiteral(((IntPtg) t).getValue())),
                new WhatIf(p, lessEqualPtg, t -> parseLeq()),
                new WhatIf(p, lessThanPtg, t -> parseLt()),
                new WhatIf(p, memErrPtg, (Ptg t) -> parseMemErrPtg((MemErrPtg) t)),
                new WhatIf(p, missingArgPtg, (Ptg t) -> parseMissingArgPtg(row, column)),
                new WhatIf(p, multiplyPtg, t -> parseMult()),
                new WhatIf(p, namePtg, (Ptg t) -> parseNamePtg((NamePtg) t)),
                new WhatIf(p, notEqualPtg, t -> parseNeq()),
                new WhatIf(p, numberPtg, t -> parseFloatLiteral(((NumberPtg) t).getValue())),
                new WhatIf(p, parenthesisPtg, t -> parseParenthesisFormula()),
                new WhatIf(p, percentPtg, t -> percentFormula()),
                new WhatIf(p, powerPtg, t -> parsePower()),
                new WhatIf(p, ref3DPxg, (Ptg t) -> parseRef3DPxg((Ref3DPxg) t)),
                new WhatIf(p, refErrorPtg, (Ptg t) -> parseReferenceErrorLiteral()),
                new WhatIf(p, refPtg, (Ptg t) -> parseRefPtg((RefPtg) t)),
                new WhatIf(p, stringPtg, (Ptg t) -> parseStringLiteral(((StringPtg) t).getValue())),
                new WhatIf(p, subtractPtg, t -> parseSub()),
                new WhatIf(p, unaryMinusPtg, (Ptg t) -> parseMinus()),
                new WhatIf(p, unaryPlusPtg, (Ptg t) -> parsePlus()),
                new WhatIf(p, unionPtg, t -> parseUnion()),
                new WhatIf(p, unknownPtg, (Ptg t) -> parseUnknownPtg((UnknownPtg) t))
        )) {
            stream.filter((WhatIf t) -> t.predicate.test(t.ptg)).forEach(t -> t.consumer.accept(t.ptg));
        } catch (Exception e) {
            err.println("parse: " + p.getClass().getSimpleName());
            err.println(this.sheetName + "row:" + row + "column:" + column + e.getMessage());
            e.printStackTrace();
        }
    }

    private void parseMemErrPtg(@NotNull MemErrPtg t) {
        err("MemErrPtg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseDeleted3DPxg(@NotNull Deleted3DPxg t) {
        err("Deleted3DPxg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseDeletedRef3DPtg(@NotNull DeletedRef3DPtg t) {
        err("DeletedRef3DPtg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseMissingArgPtg(int row, int column) {
        parseMissingArguments(row, column);
    }

    private void parseDeletedArea3DPtg(@NotNull DeletedArea3DPtg t) {
        err("DeletedArea3DPtg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseAreaErrPtg(@NotNull AreaErrPtg t) {
        err("AreaErrPtg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseUnknownPtg(@NotNull UnknownPtg t) {
        err("Error Unknown Ptg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseArea3DPxg(@NotNull Area3DPxg t) {
        // Area3DPxg is XSSF Area 3D Reference (Sheet + Area) Defined an area in an
        // external or different sheet.
        // This is XSSF only, as it stores the sheet / book references in String
        // form. The HSSF equivalent using indexes is Area3DPtg
        String sheetName = t.getSheetName();
        int sheetIndex = helper.getSheetIndex(sheetName);
        SHEET tSHEET = new SHEET(sheetName, sheetIndex);
        String area = helper.getArea(t);
        parseArea3D(helper.getRANGE(sheetName, t), tSHEET, area);
    }

    private void parseArea3D(RANGE tRANGE, @NotNull SHEET tSHEET, String area) {
        //Sheet2!A1:B1 (Sheet + AREA/RANGE)
        var term = new PrefixReferenceItem(tSHEET, area, tRANGE);
        term.setSheetIndex(tSHEET.getIndex());
        term.setSheetName(tSHEET.getName());
        unordered.add(term);
        stack.push(term);
    }

    private void parseRef3DPxg(@NotNull Ref3DPxg t) {
        //Title: XSSF 3D Reference
        //Description: Defines a cell in an external or different sheet.
        //REFERENCE:
        //This is XSSF only, as it stores the sheet / book references in String form. The HSSF equivalent using indexes is Ref3DPtg
        int extWorkbookNumber = t.getExternalWorkbookNumber();
        String sheetName = t.getSheetName();
        int sheetIndex = helper.getSheetIndex(sheetName);
        SHEET tSHEET = new SHEET(sheetName, sheetIndex);
        FILE tFILE = new FILE(extWorkbookNumber, tSHEET);
        String cellref = helper.getCellRef(t);
        if ( this.sheetIndex != sheetIndex ) {
            Sheet extSheet = this.book.getSheet(sheetName);
            if ( extSheet != null ) {
                CellReference cr = new CellReference(cellref);
                Row row = extSheet.getRow(cr.getRow());
                Cell cell = row.getCell(cr.getCol());
                this.ext.add(cell);
                verbose("Loosing!!! reference[ext] " + tSHEET.toString() + "" + cellref);
            }
        }
        if ( extWorkbookNumber > 0 ) parseReference(tFILE, cellref);
        else parseReference(tSHEET, cellref);
    }

    private void parseReference(SHEET tSHEET, String cellref) {
        var term = new PrefixReferenceItem(tSHEET, cellref, null);
        setOwnProperty(term);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseAreaPtg(@NotNull AreaPtg t) {
        parseRangeReference(helper.getRANGE(sheet, t));
    }

    private void parseRangeReference(RANGE tRANGE) {
        var rangeReference = new RangeReference(tRANGE.getFirst(), tRANGE.getLast());
        setOwnProperty(rangeReference);
        rangeReference.setAsArea();//is area not a cell with ref to area
        rangeReference.add(tRANGE.values());
        graph.addNode(rangeReference);
        stack.push(rangeReference);
    }

    private void parseNamePtg(@NotNull NamePtg t) {
        RangeInternal range = null;
        Ptg[] ptgs = helper.getName(t);
        String name = helper.getNameText(t);
        int sheetIndex = 0;
        for (Ptg ptg : ptgs) {
            if ( ptg != null ) {
                if ( ptg instanceof Area3DPxg ) {
                    Area3DPxg area3DPxg = (Area3DPxg) ptg;
                    range = new RangeInternal(book, area3DPxg.getSheetName(), area3DPxg);
                    sheetIndex = helper.getSheetIndex(area3DPxg.getSheetName());
                }
            }
        }
        RANGE tRANGE = Objects.requireNonNull(range).getRANGE();
        NamedRange term = new NamedRange(name, tRANGE);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(range.getSheetName());
        parseNamedRange(term);
    }

    private void parseNamedRange(NamedRange tNamedRange) {
        stack.push(tNamedRange);
    }

    private void parseRefPtg(@NotNull RefPtg t) {
        Row rowObject = sheet.getRow(t.getRow());
        Object value = null;
        if ( rowObject != null ) {
            Cell c = rowObject.getCell(t.getColumn());
            value = Helper.valueOf(c);
        }
        CELL cellRef = new CELL(t.getRow(), t.getColumn());
        cellRef.setValue(value);
        parseCELL_REFERENCE(cellRef);
    }

    private void parseCELL_REFERENCE(@NotNull CELL tCELL_REFERENCE) {
        setOwnProperty(tCELL_REFERENCE);
        this.unordered.add(tCELL_REFERENCE);
        stack.push(tCELL_REFERENCE);
    }

    private void parseArrayPtg(@NotNull ArrayPtg t) {
        parseConstantArray(t.getTokenArrayValues());
    }

    private void parseConstantArray(Object[][] array) {
        var term = new ConstantArray(array);
        setOwnProperty(term);
        stack.push(term);
    }

    private void parseAttrPtg(@NotNull AttrPtg t) {
        if ( t.isSum() ) parseSum();
    }

    private void parseFuncVarPtg(@NotNull FuncVarPtg t) {
        if ( t.getNumberOfOperands() == 0 ) parseFunc(t.getName());
        else parseFunc(t.getName(), t.getNumberOfOperands());
    }

    private void parseFuncPtg(@NotNull FuncPtg t) {
        if ( t.getNumberOfOperands() == 0 ) parseFunc(t.getName());
        else parseFunc(t.getName(), t.getNumberOfOperands());
    }

    private void parseFunc(String name, int arity) {
        try {
            builtInFunction(arity, name);
        } catch (UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e, rowFormula, colFormula);
        }
    }

    //@todo impl. DATE
    private void parseErrorLiteral(ErrPtg t) {
        String text;
        if ( t == NULL_INTERSECTION ) text = "#NULL!";
        else if ( t == DIV_ZERO ) text = "#DIV/0!";
        else if ( t == VALUE_INVALID ) text = "#VALUE!";
        else if ( t == REF_INVALID ) text = "#REF!";
        else if ( t == NAME_INVALID ) text = "#NAME?";
        else if ( t == NUM_ERROR ) text = "#NUM!";
        else if ( t == N_A ) text = "#N/A";
        else text = "FIXME!";
        var term = new ERROR(text);
        parseErrorLiteral(term);
    }

    private void parseErrorLiteral(@NotNull ERROR term) {
        setOwnProperty(term);
        err(term.toString(), rowFormula, colFormula);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseBooleanLiteral(Boolean bool) {
        var term = new BOOL(bool);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseStringLiteral(String string) {
        var term = new TEXT(string);
        graph.addNode(term);
        stack.push(term);
    }


    private void parseIntLiteral(Integer value) {
        var term = new INT(value);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseFloatLiteral(Double value) {
        var term = new FLOAT(value);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseReferenceErrorLiteral() {
        //#REF
        ERROR_REF term = new ERROR_REF();
        term.setColumn(colFormula);
        term.setRow(rowFormula);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(sheetName);
        term.setSingleSheet(this.isSingleSheet);
        stack.push(term);
        err("", rowFormula, colFormula);
    }


    public void parse() {
        this.isSingleSheet = this.book.getNumberOfSheets() == 1;
        for (Sheet currentSheet : this.book) parse(currentSheet);
        verbose("** topological sorting beginning...");
        sort();
    }


    private void sort() {
        if ( unordered.singleton() ) {
            ordered = new StartList();
            ordered.add(unordered.get(0));
            return;
        }
        ordered = graph.topologicalSort();
    }


    private void parseFormula(@NotNull Start obj) {
        obj.setColumn(colFormula);
        obj.setRow(rowFormula);
        obj.setSheetIndex(sheetIndex);
        obj.setSheetName(sheetName);
        obj.setSingleSheet(this.isSingleSheet);
        unordered.add(obj);
    }

    private void setOwnProperty(Start start) {
        start.setColumn(colFormula);
        start.setRow(rowFormula);
        start.setSheetIndex(sheetIndex);
        start.setSheetName(sheetName);
        start.setSingleSheet(this.isSingleSheet);
    }






    // TERMINAL AND NON TERMINAL BEGIN


    private void parseParenthesisFormula() {
        var formula = (Formula) stack.pop();
        var parFormula = new ParenthesisFormula(formula);
        setOwnProperty(parFormula);
        stack.push(parFormula);
    }


    private void parseUDF(String arguments) {
        var term = new UDF(arguments);
        setOwnProperty(term);
        unordered.add(term);
        stack.push(term);
    }

    /**
     * F=F
     */

    private void parseEq() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var eq = new Eq(lFormula, rFormula);
        setOwnProperty(eq);
        graph.add(eq);
        stack.push(eq);
    }

    /**
     * F<F
     */

    private void parseLt() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var lt = new Lt(lFormula, rFormula);
        setOwnProperty(lt);
        graph.add(lt);
        stack.push(lt);
    }

    private void parseGt() {
        // F>F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var gt = new Gt(lFormula, rFormula);
        setOwnProperty(gt);
        graph.add(gt);
        stack.push(gt);
    }

    private void parseLeq() {
        // F<=F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var leq = new Leq(lFormula, rFormula);
        setOwnProperty(leq);
        graph.add(leq);
        stack.push(leq);
    }

    private void parseGteq() {
        // F>=F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var gteq = new GtEq(lFormula, rFormula);
        setOwnProperty(gteq);
        graph.add(gteq);
        stack.push(gteq);
    }

    private void parseNeq() {
        // F<>F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var neq = new Neq(lFormula, rFormula);
        setOwnProperty(neq);
        graph.add(neq);
        stack.push(neq);
    }

    private void parseConcat() {
        // F&F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var concat = new Concat(lFormula, rFormula);
        setOwnProperty(concat);
        graph.add(concat);
        stack.push(concat);
    }

    private void parseAdd() {
        // F+F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var add = new Add(lFormula, rFormula);
        setOwnProperty(add);
        graph.add(add);
        stack.push(add);
    }

    private void parseSub() {
        // F-F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var sub = new Sub(lFormula, rFormula);
        setOwnProperty(sub);
        graph.add(sub);
        stack.push(sub);
    }

    private void parseMult() {
        // F*F
        if ( stack.empty() ) return;
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var mult = new Mult(lFormula, rFormula);
        setOwnProperty(mult);
        graph.add(mult);
        stack.push(mult);
    }

    private void parseDiv() {
        // F/F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var div = new Divide(lFormula, rFormula);
        setOwnProperty(div);
        graph.add(div);
        stack.push(div);
    }

    private void parsePower() {
        // F^F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var power = new Power(lFormula, rFormula);
        setOwnProperty(power);
        graph.add(power);
        stack.push(power);
    }

    private void percentFormula() {
        // F%
        var formula = (Formula) stack.pop();
        var percentFormula = new PercentFormula(formula);
        setOwnProperty(percentFormula);
        graph.addNode(percentFormula);
        stack.push(percentFormula);
    }

    private void parseCELLlinked(@NotNull CELL tCELL_REFERENCE) {
        setOwnProperty(tCELL_REFERENCE);
        this.unordered.add(tCELL_REFERENCE);
        stack.push(tCELL_REFERENCE);
        graph.addNode(tCELL_REFERENCE);
    }

    private void parseReference(FILE tFILE, String cellref) {
        // Used
        // Sheet2!A1 (Sheet + parseCELL_REFERENCE)
        // External references: External references are normally in the form [File]Sheet!Cell
        var term = new PrefixReferenceItem(tFILE, cellref, null);
        setOwnProperty(term);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseSum() {
        // SUM(Arguments)
        var args = stack.pop();
        if ( args instanceof Reference || args instanceof OFFSET ) {
            args.setSheetIndex(sheetIndex);
            args.setSheetName(sheetName);
            args.setAsArea();
            unordered.add(args);
        } else {
            err("Not RangeReference " + args.getClass().getSimpleName() + " " + args.toString(), rowFormula, colFormula);
        }
        var term = new SUM((Formula) args);
        setOwnProperty(term);
        unordered.add(term);
        graph.add(term);
        stack.push(term);
    }

    private void parseFunc(String name) {
        try {
            builtinFunction(name);
        } catch (UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e, rowFormula, colFormula);
        }
    }

    private void parsePlus() {
        // +
        var formula = (Formula) stack.pop();
        var plus = new Plus(formula);
        plus.setSheetName(sheetName);
        plus.setSheetIndex(sheetIndex);
        graph.addNode(plus);
        stack.push(plus);
    }

    private void parseMinus() {
        // -
        var formula = (Formula) stack.pop();
        var minus = new Minus(formula);
        setOwnProperty(minus);
        graph.addNode(minus);
        stack.push(minus);
    }


    private void builtInFunction(int arity, String name) throws UnsupportedBuiltinException {
        var factory = new BuiltinFactory();
        factory.create(arity, name);
        var builtinFunction = (EXCEL_FUNCTION) factory.getBuiltInFunction();
        Start[] args = factory.getArgs();
        for (int i = arity - 1; i >= 0; i--) if ( !stack.empty() ) args[i] = stack.pop();

        setOwnProperty(builtinFunction);
        graph.addNode(builtinFunction);
        for (Start arg : args) {
            if ( arg instanceof RangeReference /*|| arg instanceof CELL*/ || arg instanceof PrefixReferenceItem || arg instanceof ReferenceItem ) {
                if ( unordered.add(arg) ) {
                    graph.addNode(arg);
                    graph.addEdge(arg, builtinFunction);
                }
            }
        }
        stack.push(builtinFunction);
    }

    private void builtinFunction(String name) throws UnsupportedBuiltinException {
        var factory = new BuiltinFactory();
        factory.create(0, name);
        var builtinFunction = (EXCEL_FUNCTION) factory.getBuiltInFunction();
        stack.push(builtinFunction);
    }

    public StartList getList() {
        return ordered;
    }

    private void parseIntersection() {
        //F F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var intersection = new Intersection(lFormula, rFormula);
        setOwnProperty(intersection);
        graph.add(intersection);
        stack.push(intersection);
    }

    private void parseUnion() {
        //F,F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var union = new Union(lFormula, rFormula);
        setOwnProperty(union);
        graph.add(union);
        stack.push(union);
    }



    private void parseMissingArguments(int row, int column) {
        err("Missing ExcelFunction Arguments for cell: " + Start.cellAddress(row, column, sheetName), row, column);
    }

    private void doesFormulaReferToDeletedCell(int row, int column) {
        err(Start.cellAddress(row, column, sheetName) + " does formula refer to deleted cell", row, column);
    }

    private void err(String string, int row, int column) {
        err.println(Start.cellAddress(row, column, sheetName) + " parseErrorLiteral: " + string);
    }

    class WhatIf {

        final Ptg ptg;
        final Predicate<Ptg> predicate;
        final Consumer<Ptg> consumer;

        WhatIf(Ptg ptg, Predicate<Ptg> predicate, Consumer<Ptg> consumer) {
            this.ptg = ptg;
            this.predicate = predicate;
            this.consumer = consumer;
        }
    }

    class RangeInternal {

        @NotNull
        private final RANGE tRANGE;
        private final String sheetName;


        RangeInternal(Workbook workbook, String sheetnamne, Area3DPxg t) {
            Helper helper = new Helper(workbook);
            int firstRow = t.getFirstRow();
            int firstColumn = t.getFirstColumn();
            sheetName = sheetnamne;
            int lastRow = t.getLastRow();
            int lastColumn = t.getLastColumn();

            CELL first = new CELL(firstRow, firstColumn);
            CELL last = new CELL(lastRow, lastColumn);
            tRANGE = new RANGE(first, last);
            String refs = tRANGE.toString();
            SpreadsheetVersion SPREADSHEET_VERSION = SpreadsheetVersion.EXCEL2007;
            AreaReference area = new AreaReference(sheetnamne + "!" + refs, SPREADSHEET_VERSION);
            List<Cell> cells = helper.fromRange(area);

            for (Cell cell : cells)
                if ( cell != null ) {
                    tRANGE.add(Helper.valueOf(cell));
                }
        }


        String getSheetName() {
            return sheetName;
        }

        @NotNull
        RANGE getRANGE() {
            return tRANGE;
        }

    }

}
