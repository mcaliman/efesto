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

    private boolean verbose = false;

    private int column;//Current Formula Column
    private int row;//Current Formula Row

    private SHEET cSHEET;//current sheet

    private boolean singleSheet;//is single sheet or not?
    private Workbook book;
    private Helper helper;
    private List<Cell> ext;
    private int counterFormulas;//formula counters
    private Sheet sheet;//(Work)Sheet
    private String fileName;

    private StartList unordered;
    private StartList ordered;
    private StartGraph graph;
    private Stack<Start> stack;

    public Parser(String filename) throws IOException, InvalidFormatException {
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


    private void parse(Sheet sheet) {
        this.sheet = sheet;
        int index = getSheetIndex();
        String name = getSheetName();
        this.cSHEET = new SHEET(name, index);
        verbose("Parsing sheet-name:" + cSHEET.getName());
        for (Row row : sheet)
            for (Cell cell : row)
                if ( cell != null ) parse(cell);
                else err("Cell is null.");
    }


    private void parse(Cell cell) {
        if ( cell.getCellType() == CELL_TYPE_FORMULA ) {
            parseFormula(cell);
            this.counterFormulas++;
        } else if ( this.ext.contains(cell) ) {
            verbose("Recover loosed cell!");
            Object value = Helper.valueOf(cell);
            CELL elem = new CELL(cell.getRowIndex(), cell.getColumnIndex());
            elem.setValue(value);
            String name = getSheetName(cell);
            int index = getSheetIndex(cell);
            elem.setSheetName(name);
            elem.setSheetIndex(index);
            parseCELLlinked(elem);
            this.ext.remove(cell);
        } else if ( !this.ext.contains(cell) && !isCellEmpty(cell) ) {
            //Non è formula non è nelle celle utili collezionate
            out.println("Cella di interesse? " + cell.toString());

        }
    }

    private void parseFormula(Cell cell) {
        verbose("Cell:" + cell.getClass().getSimpleName() + " " + cell.toString() + " " + cell.getCellType());
        this.column = cell.getColumnIndex();
        this.row = cell.getRowIndex();
        String formulaAddress = getCellAddress();
        Ptg[] formulaPtgs = helper.tokens(this.sheet, this.row, this.column);
        if ( formulaPtgs == null ) {
            String formulaText = cell.getCellFormula();
            err("ptgs empty or null for address " + formulaAddress);
            parseUDF(formulaText);
            return;
        }
        Start start = parse(formulaPtgs);
        if ( Objects.nonNull(start) ) {
            start.setSingleSheet(this.singleSheet);
            parseFormula(start);
        }
    }

    private Start parse(Ptg[] ptgs) {
        stack.empty();
        if ( Ptg.doesFormulaReferToDeletedCell(ptgs) ) doesFormulaReferToDeletedCell();
        for (Ptg ptg : ptgs) parse(ptg);
        Start start = null;
        if ( !stack.empty() ) start = stack.pop();
        return start;
    }

    private void parseUDF(String arguments) {
        var elem = new UDF(arguments);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.cSHEET.getIndex());
        elem.setSheetName(this.cSHEET.getName());
        elem.setSingleSheet(this.singleSheet);
        unordered.add(elem);
        stack.push(elem);
    }

    private void parse(Ptg p) {
        verbose("parse: " + p.getClass().getSimpleName());
        try (Stream<WhatIf> stream = Stream.of(
                new WhatIf(p, arrayPtg, (Ptg t) -> parseConstantArray((ArrayPtg) t)),
                new WhatIf(p, addPtg, (Ptg t) -> parseAdd()),
                new WhatIf(p, area3DPxg, (Ptg t) -> parseArea3DPxg((Area3DPxg) t)),
                new WhatIf(p, areaErrPtg, (Ptg t) -> parseErrPtg(t)),
                new WhatIf(p, areaPtg, (Ptg t) -> parseAreaPtg((AreaPtg) t)),
                new WhatIf(p, attrPtg, (Ptg t) -> parseAttrPtg((AttrPtg) t)),
                new WhatIf(p, boolPtg, t -> parseBOOL(((BoolPtg) t).getValue())),
                new WhatIf(p, concatPtg, t -> parseConcat()),
                new WhatIf(p, deleted3DPxg, (Ptg t) -> parseErrPtg(t)),
                new WhatIf(p, deletedArea3DPtg, (Ptg t) -> parseErrPtg(t)),
                new WhatIf(p, deletedRef3DPtg, (Ptg t) -> parseErrPtg(t)),
                new WhatIf(p, dividePtg, t -> parseDiv()),
                new WhatIf(p, equalPtg, t -> parseEq()),
                new WhatIf(p, errPtg, (Ptg t) -> parseERROR((ErrPtg) t)),
                new WhatIf(p, funcPtg, (Ptg t) -> parseFuncPtg((FuncPtg) t)),
                new WhatIf(p, funcVarPtg, (Ptg t) -> parseFuncVarPtg((FuncVarPtg) t)),
                new WhatIf(p, greaterEqualPtg, t -> parseGteq()),
                new WhatIf(p, greaterThanPtg, t -> parseGt()),
                new WhatIf(p, intersectionPtg, t -> parseIntersection()),
                new WhatIf(p, intPtg, t -> parseINT(((IntPtg) t).getValue())),
                new WhatIf(p, lessEqualPtg, t -> parseLeq()),
                new WhatIf(p, lessThanPtg, t -> parseLt()),
                new WhatIf(p, memErrPtg, (Ptg t) -> parseErrPtg(t)),
                new WhatIf(p, missingArgPtg, (Ptg t) -> parseMissingArguments()),
                new WhatIf(p, multiplyPtg, t -> parseMult()),
                new WhatIf(p, namePtg, (Ptg t) -> parseNamedRange((NamePtg) t)),
                new WhatIf(p, notEqualPtg, t -> parseNeq()),
                new WhatIf(p, numberPtg, t -> parseFLOAT(((NumberPtg) t).getValue())),
                new WhatIf(p, parenthesisPtg, t -> parseParenthesisFormula()),
                new WhatIf(p, percentPtg, t -> percentFormula()),
                new WhatIf(p, powerPtg, t -> parsePower()),
                new WhatIf(p, ref3DPxg, (Ptg t) -> parseRef3DPxg((Ref3DPxg) t)),
                new WhatIf(p, refErrorPtg, (Ptg t) -> parseERRORREF()),
                new WhatIf(p, refPtg, (Ptg t) -> parseRefPtg((RefPtg) t)),
                new WhatIf(p, stringPtg, (Ptg t) -> parseTEXT(((StringPtg) t).getValue())),
                new WhatIf(p, subtractPtg, t -> parseSub()),
                new WhatIf(p, unaryMinusPtg, (Ptg t) -> parseMinus()),
                new WhatIf(p, unaryPlusPtg, (Ptg t) -> parsePlus()),
                new WhatIf(p, unionPtg, t -> parseUnion()),
                new WhatIf(p, unknownPtg, (Ptg t) -> parseErrPtg(t))
        )) {
            stream.filter((WhatIf t) -> t.predicate.test(t.ptg)).forEach(t -> t.consumer.accept(t.ptg));
        } catch (Exception e) {
            err.println("parse: " + p.getClass().getSimpleName() + " " + this.cSHEET.getName() + "row:" + row + "column:" + column + e.getMessage());
            e.printStackTrace();
        }
    }


    private void parseErrPtg(Ptg t) {
        err(t.getClass().getName() + ": " + t.toString());
    }

    private void parseMissingArguments() {
        err("Missing ExcelFunction Arguments for cell: " + getCellAddress());
    }

    private void parseArea3DPxg(Area3DPxg t) {
        // Area3DPxg is XSSF Area 3D Reference (Sheet + Area) Defined an area in an
        // external or different sheet.
        // This is XSSF only, as it stores the sheet / book references in String
        // form. The HSSF equivalent using indexes is Area3DPtg
        String name = t.getSheetName();
        int index = helper.getSheetIndex(name);
        SHEET tSHEET = new SHEET(name, index);
        String area = helper.getArea(t);
        parseArea3D(helper.getRANGE(name, t), tSHEET, area);
    }

    private void parseArea3D(RANGE tRANGE, SHEET tSHEET, String area) {
        //Sheet2!A1:B1 (Sheet + AREA/RANGE)
        var elem = new PrefixReferenceItem(tSHEET, area, tRANGE);
        elem.setSheetIndex(tSHEET.getIndex());
        elem.setSheetName(tSHEET.getName());
        unordered.add(elem);
        stack.push(elem);
    }

    private void parseRef3DPxg(Ref3DPxg t) {
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
        if ( this.cSHEET.getIndex() != sheetIndex ) {
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
        var elem = new PrefixReferenceItem(tSHEET, cellref, null);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.addNode(elem);
        stack.push(elem);
    }

    private void parseAreaPtg(AreaPtg t) {
        RANGE tRANGE = helper.getRANGE(sheet, t);
        // RangeReference
        var elem = new RangeReference(tRANGE.getFirst(), tRANGE.getLast());
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);

        elem.setAsArea();//is area not a cell with ref to area
        elem.add(tRANGE.values());
        graph.addNode(elem);
        stack.push(elem);

    }

    private void parseNamedRange(@NotNull NamePtg t) {
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
        NamedRange elem = new NamedRange(name, tRANGE);
        elem.setSheetIndex(sheetIndex);
        elem.setSheetName(range.getSheetName());
        stack.push(elem);
    }

    private void parseRefPtg(@NotNull RefPtg t) {
        Row rowObject = sheet.getRow(t.getRow());
        Object value = null;
        if ( rowObject != null ) {
            Cell c = rowObject.getCell(t.getColumn());
            value = Helper.valueOf(c);
        }
        CELL elem = new CELL(t.getRow(), t.getColumn());
        elem.setValue(value);
        //parse CELL
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        this.unordered.add(elem);
        stack.push(elem);
    }


    private void parseConstantArray(@NotNull ArrayPtg t) {
        Object[][] array = t.getTokenArrayValues();
        // ConstantArray
        var elem = new ConstantArray(array);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        stack.push(elem);
    }

    private void parseAttrPtg(@NotNull AttrPtg t) {
        if ( t.isSum() ) parseSum();
    }

    private void parseSum() {
        // SUM(Arguments)
        var args = stack.pop();
        if ( args instanceof Reference || args instanceof OFFSET ) {
            args.setSHEET(cSHEET);
            args.setAsArea();
            unordered.add(args);
        } else {
            err("Not RangeReference " + args.getClass().getSimpleName() + " " + args.toString());
        }
        var elem = new SUM((Formula) args);

        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(cSHEET.getIndex());
        elem.setSheetName(cSHEET.getName());
        elem.setSingleSheet(this.singleSheet);

        unordered.add(elem);
        graph.add(elem);
        stack.push(elem);
    }

    private void parseFuncVarPtg(@NotNull FuncVarPtg t) {
        int arity = t.getNumberOfOperands();
        String name = t.getName();
        if ( arity == 0 ) parseFunc(name);
        else parseFunc(name, arity);
    }

    private void parseFuncPtg(@NotNull FuncPtg t) {
        int arity = t.getNumberOfOperands();
        String name = t.getName();
        if ( arity == 0 ) parseFunc(name);
        else parseFunc(name, arity);
    }

    private void parseFunc(String name, int arity) {
        try {
            builtInFunction(arity, name);
        } catch (UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e);
        }
    }

    //@todo impl. DATE
    private void parseERROR(ErrPtg t) {
        String text;
        if ( t == NULL_INTERSECTION ) text = "#NULL!";
        else if ( t == DIV_ZERO ) text = "#DIV/0!";
        else if ( t == VALUE_INVALID ) text = "#VALUE!";
        else if ( t == REF_INVALID ) text = "#REF!";
        else if ( t == NAME_INVALID ) text = "#NAME?";
        else if ( t == NUM_ERROR ) text = "#NUM!";
        else if ( t == N_A ) text = "#N/A";
        else text = "FIXME!";


        // ERROR
        var elem = new ERROR(text);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(cSHEET.getIndex());
        elem.setSheetName(cSHEET.getName());
        elem.setSingleSheet(this.singleSheet);

        err(elem.toString());
        graph.addNode(elem);
        stack.push(elem);
    }

    private void parseBOOL(Boolean bool) {
        var elem = new BOOL(bool);
        graph.addNode(elem);
        stack.push(elem);
    }

    private void parseTEXT(String string) {
        var elem = new TEXT(string);
        graph.addNode(elem);
        stack.push(elem);
    }

    private void parseINT(Integer value) {
        var elem = new INT(value);
        graph.addNode(elem);
        stack.push(elem);
    }

    private void parseFLOAT(Double value) {
        var elem = new FLOAT(value);
        graph.addNode(elem);
        stack.push(elem);
    }

    private void parseERRORREF() {
        //#REF
        ERRORREF elem = new ERRORREF();
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(cSHEET.getIndex());
        elem.setSheetName(cSHEET.getName());
        elem.setSingleSheet(this.singleSheet);
        stack.push(elem);
        err("");
    }

    public void parse() {
        this.singleSheet = this.book.getNumberOfSheets() == 1;
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

    private void parseFormula(@NotNull Start elem) {
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        unordered.add(elem);
    }
    private void parseParenthesisFormula() {
        var formula = (Formula) stack.pop();
        var elem = new ParenthesisFormula(formula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        stack.push(elem);
    }

    private void parseEq() {
        // F=F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Eq(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parseLt() {
        // F<F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Lt(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parseGt() {
        // F>F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Gt(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parseLeq() {
        // F<=F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Leq(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parseGteq() {
        // F>=F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new GtEq(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parseNeq() {
        // F<>F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Neq(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parseConcat() {
        // F&F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Concat(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parseAdd() {
        // F+F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Add(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parseSub() {
        // F-F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Sub(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parseMult() {
        // F*F
        if ( stack.empty() ) return;
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Mult(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parseDiv() {
        // F/F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Divide(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void parsePower() {
        // F^F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Power(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    private void percentFormula() {
        // F%
        var formula = (Formula) stack.pop();
        var elem = new PercentFormula(formula);
        elem.setColumn(column);
        elem.setRow(row);

        elem.setSHEET(cSHEET);

        elem.setSingleSheet(this.singleSheet);
        graph.addNode(elem);
        stack.push(elem);
    }

    private void parseCELLlinked(@NotNull CELL elem) {
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        this.unordered.add(elem);
        stack.push(elem);
        graph.addNode(elem);
    }

    private void parseReference(FILE tFILE, String cellref) {
        // Used
        // Sheet2!A1 (Sheet + parseCELL_REFERENCE)
        // External references: External references are normally in the form [File]Sheet!Cell
        var elem = new PrefixReferenceItem(tFILE, cellref, null);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.addNode(elem);
        stack.push(elem);
    }


    private void parseFunc(String name) {
        try {
            builtinFunction(name);
        } catch (UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e);
        }
    }

    private void parsePlus() {
        // +
        var formula = (Formula) stack.pop();
        var elem = new Plus(formula);
        elem.setSHEET(cSHEET);
        graph.addNode(elem);
        stack.push(elem);
    }
    private void parseMinus() {
        // -
        var formula = (Formula) stack.pop();
        var elem = new Minus(formula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.addNode(elem);
        stack.push(elem);
    }


    private void builtInFunction(int arity, String name) throws UnsupportedBuiltinException {
        var factory = new BuiltinFactory();
        factory.create(arity, name);
        var builtinFunction = (EXCEL_FUNCTION) factory.getBuiltInFunction();
        Start[] args = factory.getArgs();
        for (int i = arity - 1; i >= 0; i--) if ( !stack.empty() ) args[i] = stack.pop();

        builtinFunction.setColumn(column);
        builtinFunction.setRow(row);
        builtinFunction.setSHEET(cSHEET);
        builtinFunction.setSingleSheet(this.singleSheet);

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
        var rFormula = (Formula) this.stack.pop();
        var lFormula = (Formula) this.stack.pop();
        var elem = new Intersection(lFormula, rFormula);
        elem.setColumn(this.column);
        elem.setRow(this.row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        this.graph.add(elem);
        this.stack.push(elem);
    }

    private void parseUnion() {
        //F,F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Union(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSHEET(cSHEET);
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }


    private void doesFormulaReferToDeletedCell() {
        err(getCellAddress() + " does formula refer to deleted cell");
    }

    private void err(String string) {
        err.println(getCellAddress() + " error: " + string);
        //throw new RuntimeException(getCellAddress() + " error: " + string);
    }

    private String getCellAddress() {
        return Start.cellAddress(this.row, this.column, this.cSHEET.getName());
    }

    public void setVerbose(boolean verbose) {
        this.verbose = verbose;
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

    private String getSheetName(Cell cell) {
        return cell.getSheet().getSheetName();
    }

    private int getSheetIndex(Cell cell) {
        return helper.getSheetIndex(cell.getSheet().getSheetName());
    }

    private int getSheetIndex() {
        return this.book.getSheetIndex(sheet);
    }

    private String getSheetName() {
        return this.sheet.getSheetName();
    }

    private boolean isCellEmpty(final Cell cell) {
        if ( cell == null ) { // use row.getCell(x, Row.CREATE_NULL_AS_BLANK) to avoid null cells
            return true;
        }

        if ( cell.getCellType() == Cell.CELL_TYPE_BLANK ) {
            return true;
        }

        if ( cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().trim().isEmpty() ) {
            return true;
        }

        return false;
    }

    // INNER CLASS


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
