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
    private boolean verbose = false;

    private int formulaColumn;//Current Formula Column
    private int formulaRow;//Current Formula Row
    private int sheetIndex;//Current Sheet Index
    private String sheetName;//Current Sheet Name

    private boolean isSingleSheet;
    private Workbook book;
    private Helper helper;
    private List<Cell> ext;
    private int counterFormulas;
    private Sheet sheet;//(Work)Sheet
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

    private void parse(@NotNull Sheet sheet) {
        this.sheet = sheet;
        this.sheetIndex = book.getSheetIndex(sheet);
        this.sheetName = sheet.getSheetName();
        verbose("Parsing sheet-name:" + this.sheetName);
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
            Object obj = Helper.valueOf(cell);
            CELL cellRef = new CELL(cell.getRowIndex(), cell.getColumnIndex());
            cellRef.setValue(obj);
            cellRef.setSheetName(cell.getSheet().getSheetName());
            cellRef.setSheetIndex(helper.getSheetIndex(cell.getSheet().getSheetName()));
            parseCELLlinked(cellRef);
            this.ext.remove(cell);
        } else if ( !this.ext.contains(cell) && !isCellEmpty(cell) ) {
            //Non è formula non è nelle celle utili collezionate
            out.println("Cella di interesse? " + cell.toString());

        }
    }

    private void parseFormula(Cell cell) {
        verbose("Cell:" + cell.getClass().getSimpleName() + " " + cell.toString() + " " + cell.getCellType());
        formulaColumn = cell.getColumnIndex();
        formulaRow = cell.getRowIndex();
        String formulaAddress = getCellAddress();
        Ptg[] formulaPtgs = helper.tokens(this.sheet, this.formulaRow, this.formulaColumn);
        if ( formulaPtgs == null ) {
            String formulaText = cell.getCellFormula();
            err("ptgs empty or null for address " + formulaAddress);
            parseUDF(formulaText);
            return;
        }
        Start start = parse(formulaPtgs);
        if ( Objects.nonNull(start) ) {
            start.setSingleSheet(this.isSingleSheet);
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
        var udf = new UDF(arguments);
        udf.setColumn(formulaColumn);
        udf.setRow(formulaRow);
        udf.setSheetIndex(sheetIndex);
        udf.setSheetName(sheetName);
        udf.setSingleSheet(this.isSingleSheet);
        unordered.add(udf);
        stack.push(udf);
    }

    private void parse(Ptg p) {
        verbose("parse: " + p.getClass().getSimpleName());
        try (Stream<WhatIf> stream = Stream.of(
                new WhatIf(p, arrayPtg, (Ptg t) -> parseConstantArray((ArrayPtg) t)),
                new WhatIf(p, addPtg, (Ptg t) -> parseAdd()),
                new WhatIf(p, area3DPxg, (Ptg t) -> parseArea3DPxg((Area3DPxg) t)),
                new WhatIf(p, areaErrPtg, (Ptg t) -> parseAreaErrPtg((AreaErrPtg) t)),
                new WhatIf(p, areaPtg, (Ptg t) -> parseAreaPtg((AreaPtg) t)),
                new WhatIf(p, attrPtg, (Ptg t) -> parseAttrPtg((AttrPtg) t)),
                new WhatIf(p, boolPtg, t -> parseBOOL(((BoolPtg) t).getValue())),
                new WhatIf(p, concatPtg, t -> parseConcat()),
                new WhatIf(p, deleted3DPxg, (Ptg t) -> parseDeleted3DPxg((Deleted3DPxg) t)),
                new WhatIf(p, deletedArea3DPtg, (Ptg t) -> parseDeletedArea3DPtg((DeletedArea3DPtg) t)),
                new WhatIf(p, deletedRef3DPtg, (Ptg t) -> parseDeletedRef3DPtg((DeletedRef3DPtg) t)),
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
                new WhatIf(p, memErrPtg, (Ptg t) -> parseErrPtg((MemErrPtg) t)),
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
                new WhatIf(p, unknownPtg, (Ptg t) -> parseUnknownPtg((UnknownPtg) t))
        )) {
            stream.filter((WhatIf t) -> t.predicate.test(t.ptg)).forEach(t -> t.consumer.accept(t.ptg));
        } catch (Exception e) {
            err.println("parse: " + p.getClass().getSimpleName() + " " + this.sheetName + "row:" + formulaRow + "column:" + formulaColumn + e.getMessage());
            e.printStackTrace();
        }
    }


    private void parseErrPtg(Ptg t) {
        err(t.getClass().getName() + ": " + t.toString());
    }

    /*private void parseMemErrPtg(MemErrPtg t) {
        err("MemErrPtg: " + t.toString());
    }*/

    private void parseDeleted3DPxg(Deleted3DPxg t) {
        err("Deleted3DPxg: " + t.toString());
    }

    private void parseDeletedRef3DPtg(DeletedRef3DPtg t) {
        err("DeletedRef3DPtg: " + t.toString());
    }

    private void parseMissingArguments() {
        err("Missing ExcelFunction Arguments for cell: " + getCellAddress());
    }

    private void parseDeletedArea3DPtg(DeletedArea3DPtg t) {
        err("DeletedArea3DPtg: " + t.toString());
    }

    private void parseAreaErrPtg(AreaErrPtg t) {
        err("AreaErrPtg: " + t.toString());
    }

    private void parseUnknownPtg(UnknownPtg t) {
        err("Error Unknown Ptg: " + t.toString());
    }

    private void parseArea3DPxg(Area3DPxg t) {
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

    private void parseArea3D(RANGE tRANGE, SHEET tSHEET, String area) {
        //Sheet2!A1:B1 (Sheet + AREA/RANGE)
        var term = new PrefixReferenceItem(tSHEET, area, tRANGE);
        term.setSheetIndex(tSHEET.getIndex());
        term.setSheetName(tSHEET.getName());
        unordered.add(term);
        stack.push(term);
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
        term.setColumn(formulaColumn);
        term.setRow(formulaRow);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(sheetName);
        term.setSingleSheet(this.isSingleSheet);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseAreaPtg(AreaPtg t) {
        RANGE tRANGE = helper.getRANGE(sheet, t);
        // RangeReference
        var term = new RangeReference(tRANGE.getFirst(), tRANGE.getLast());
        term.setColumn(formulaColumn);
        term.setRow(formulaRow);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(sheetName);
        term.setSingleSheet(this.isSingleSheet);

        term.setAsArea();//is area not a cell with ref to area
        term.add(tRANGE.values());
        graph.addNode(term);
        stack.push(term);

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
        NamedRange namedRange = new NamedRange(name, tRANGE);
        namedRange.setSheetIndex(sheetIndex);
        namedRange.setSheetName(range.getSheetName());
        stack.push(namedRange);
    }

    private void parseRefPtg(@NotNull RefPtg t) {
        Row rowObject = sheet.getRow(t.getRow());
        Object value = null;
        if ( rowObject != null ) {
            Cell c = rowObject.getCell(t.getColumn());
            value = Helper.valueOf(c);
        }
        CELL term = new CELL(t.getRow(), t.getColumn());
        term.setValue(value);
        //parse CELL
        term.setColumn(formulaColumn);
        term.setRow(formulaRow);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(sheetName);
        term.setSingleSheet(this.isSingleSheet);
        this.unordered.add(term);
        stack.push(term);
    }


    private void parseConstantArray(@NotNull ArrayPtg t) {
        Object[][] array = t.getTokenArrayValues();
        // ConstantArray
        var term = new ConstantArray(array);
        term.setColumn(formulaColumn);
        term.setRow(formulaRow);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(sheetName);
        term.setSingleSheet(this.isSingleSheet);
        stack.push(term);
    }

    private void parseAttrPtg(@NotNull AttrPtg t) {
        if ( t.isSum() ) parseSum();
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
            err("Not RangeReference " + args.getClass().getSimpleName() + " " + args.toString());
        }
        var term = new SUM((Formula) args);

        term.setColumn(formulaColumn);
        term.setRow(formulaRow);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(sheetName);
        term.setSingleSheet(this.isSingleSheet);

        unordered.add(term);
        graph.add(term);
        stack.push(term);
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
        var term = new ERROR(text);
        term.setColumn(formulaColumn);
        term.setRow(formulaRow);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(sheetName);
        term.setSingleSheet(this.isSingleSheet);

        err(term.toString());
        graph.addNode(term);
        stack.push(term);
    }

    private void parseBOOL(Boolean bool) {
        var term = new BOOL(bool);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseTEXT(String string) {
        var term = new TEXT(string);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseINT(Integer value) {
        var term = new INT(value);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseFLOAT(Double value) {
        var term = new FLOAT(value);
        graph.addNode(term);
        stack.push(term);
    }

    private void parseERRORREF() {
        //#REF
        ERRORREF term = new ERRORREF();
        term.setColumn(formulaColumn);
        term.setRow(formulaRow);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(sheetName);
        term.setSingleSheet(this.isSingleSheet);
        stack.push(term);
        err("");
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
    private void parseFormula(@NotNull Start formula) {
        formula.setColumn(formulaColumn);
        formula.setRow(formulaRow);
        formula.setSheetIndex(sheetIndex);
        formula.setSheetName(sheetName);
        formula.setSingleSheet(this.isSingleSheet);
        unordered.add(formula);
    }
    private void parseParenthesisFormula() {
        var formula = (Formula) stack.pop();
        var parFormula = new ParenthesisFormula(formula);
        parFormula.setColumn(formulaColumn);
        parFormula.setRow(formulaRow);
        parFormula.setSheetIndex(sheetIndex);
        parFormula.setSheetName(sheetName);
        parFormula.setSingleSheet(this.isSingleSheet);
        stack.push(parFormula);
    }

    private void parseEq() {
        // F=F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var eq = new Eq(lFormula, rFormula);
        eq.setColumn(formulaColumn);
        eq.setRow(formulaRow);
        eq.setSheetIndex(sheetIndex);
        eq.setSheetName(sheetName);
        eq.setSingleSheet(this.isSingleSheet);
        graph.add(eq);
        stack.push(eq);
    }
    private void parseLt() {
        // F<F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var lt = new Lt(lFormula, rFormula);
        lt.setColumn(formulaColumn);
        lt.setRow(formulaRow);
        lt.setSheetIndex(sheetIndex);
        lt.setSheetName(sheetName);
        lt.setSingleSheet(this.isSingleSheet);
        graph.add(lt);
        stack.push(lt);
    }
    private void parseGt() {
        // F>F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var gt = new Gt(lFormula, rFormula);
        gt.setColumn(formulaColumn);
        gt.setRow(formulaRow);
        gt.setSheetIndex(sheetIndex);
        gt.setSheetName(sheetName);
        gt.setSingleSheet(this.isSingleSheet);
        graph.add(gt);
        stack.push(gt);
    }
    private void parseLeq() {
        // F<=F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var leq = new Leq(lFormula, rFormula);
        leq.setColumn(formulaColumn);
        leq.setRow(formulaRow);
        leq.setSheetIndex(sheetIndex);
        leq.setSheetName(sheetName);
        leq.setSingleSheet(this.isSingleSheet);
        graph.add(leq);
        stack.push(leq);
    }
    private void parseGteq() {
        // F>=F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var gteq = new GtEq(lFormula, rFormula);
        gteq.setColumn(formulaColumn);
        gteq.setRow(formulaRow);
        gteq.setSheetIndex(sheetIndex);
        gteq.setSheetName(sheetName);
        gteq.setSingleSheet(this.isSingleSheet);
        graph.add(gteq);
        stack.push(gteq);
    }
    private void parseNeq() {
        // F<>F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var neq = new Neq(lFormula, rFormula);
        neq.setColumn(formulaColumn);
        neq.setRow(formulaRow);
        neq.setSheetIndex(sheetIndex);
        neq.setSheetName(sheetName);
        neq.setSingleSheet(this.isSingleSheet);
        graph.add(neq);
        stack.push(neq);
    }
    private void parseConcat() {
        // F&F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var concat = new Concat(lFormula, rFormula);
        concat.setColumn(formulaColumn);
        concat.setRow(formulaRow);
        concat.setSheetIndex(sheetIndex);
        concat.setSheetName(sheetName);
        concat.setSingleSheet(this.isSingleSheet);
        graph.add(concat);
        stack.push(concat);
    }
    private void parseAdd() {
        // F+F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var add = new Add(lFormula, rFormula);
        add.setColumn(formulaColumn);
        add.setRow(formulaRow);
        add.setSheetIndex(sheetIndex);
        add.setSheetName(sheetName);
        add.setSingleSheet(this.isSingleSheet);
        graph.add(add);
        stack.push(add);
    }
    private void parseSub() {
        // F-F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var sub = new Sub(lFormula, rFormula);
        sub.setColumn(formulaColumn);
        sub.setRow(formulaRow);
        sub.setSheetIndex(sheetIndex);
        sub.setSheetName(sheetName);
        sub.setSingleSheet(this.isSingleSheet);
        graph.add(sub);
        stack.push(sub);
    }
    private void parseMult() {
        // F*F
        if ( stack.empty() ) return;
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var mult = new Mult(lFormula, rFormula);
        mult.setColumn(formulaColumn);
        mult.setRow(formulaRow);
        mult.setSheetIndex(sheetIndex);
        mult.setSheetName(sheetName);
        mult.setSingleSheet(this.isSingleSheet);
        graph.add(mult);
        stack.push(mult);
    }
    private void parseDiv() {
        // F/F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var div = new Divide(lFormula, rFormula);
        div.setColumn(formulaColumn);
        div.setRow(formulaRow);
        div.setSheetIndex(sheetIndex);
        div.setSheetName(sheetName);
        div.setSingleSheet(this.isSingleSheet);
        graph.add(div);
        stack.push(div);
    }
    private void parsePower() {
        // F^F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var power = new Power(lFormula, rFormula);
        power.setColumn(formulaColumn);
        power.setRow(formulaRow);
        power.setSheetIndex(sheetIndex);
        power.setSheetName(sheetName);
        power.setSingleSheet(this.isSingleSheet);
        graph.add(power);
        stack.push(power);
    }
    private void percentFormula() {
        // F%
        var formula = (Formula) stack.pop();
        var percentFormula = new PercentFormula(formula);
        percentFormula.setColumn(formulaColumn);
        percentFormula.setRow(formulaRow);
        percentFormula.setSheetIndex(sheetIndex);
        percentFormula.setSheetName(sheetName);
        percentFormula.setSingleSheet(this.isSingleSheet);
        graph.addNode(percentFormula);
        stack.push(percentFormula);
    }

    private void parseCELLlinked(@NotNull CELL tCELL) {
        tCELL.setColumn(formulaColumn);
        tCELL.setRow(formulaRow);
        tCELL.setSheetIndex(sheetIndex);
        tCELL.setSheetName(sheetName);
        tCELL.setSingleSheet(this.isSingleSheet);
        this.unordered.add(tCELL);
        stack.push(tCELL);
        graph.addNode(tCELL);
    }

    private void parseReference(FILE tFILE, String cellref) {
        // Used
        // Sheet2!A1 (Sheet + parseCELL_REFERENCE)
        // External references: External references are normally in the form [File]Sheet!Cell
        var term = new PrefixReferenceItem(tFILE, cellref, null);
        term.setColumn(formulaColumn);
        term.setRow(formulaRow);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(sheetName);
        term.setSingleSheet(this.isSingleSheet);
        graph.addNode(term);
        stack.push(term);
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
        minus.setColumn(formulaColumn);
        minus.setRow(formulaRow);
        minus.setSheetIndex(sheetIndex);
        minus.setSheetName(sheetName);
        minus.setSingleSheet(this.isSingleSheet);
        graph.addNode(minus);
        stack.push(minus);
    }


    private void builtInFunction(int arity, String name) throws UnsupportedBuiltinException {
        var factory = new BuiltinFactory();
        factory.create(arity, name);
        var builtinFunction = (EXCEL_FUNCTION) factory.getBuiltInFunction();
        Start[] args = factory.getArgs();
        for (int i = arity - 1; i >= 0; i--) if ( !stack.empty() ) args[i] = stack.pop();

        builtinFunction.setColumn(formulaColumn);
        builtinFunction.setRow(formulaRow);
        builtinFunction.setSheetIndex(sheetIndex);
        builtinFunction.setSheetName(sheetName);
        builtinFunction.setSingleSheet(this.isSingleSheet);

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
        var intersection = new Intersection(lFormula, rFormula);
        intersection.setColumn(this.formulaColumn);
        intersection.setRow(this.formulaRow);
        intersection.setSheetIndex(this.sheetIndex);
        intersection.setSheetName(this.sheetName);
        intersection.setSingleSheet(this.isSingleSheet);
        this.graph.add(intersection);
        this.stack.push(intersection);
    }

    private void parseUnion() {
        //F,F
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var union = new Union(lFormula, rFormula);
        union.setColumn(formulaColumn);
        union.setRow(formulaRow);
        union.setSheetIndex(sheetIndex);
        union.setSheetName(sheetName);
        union.setSingleSheet(this.isSingleSheet);
        graph.add(union);
        stack.push(union);
    }


    private void doesFormulaReferToDeletedCell() {
        err(getCellAddress() + " does formula refer to deleted cell");
    }

    private void err(String string) {
        err.println(getCellAddress() + " error: " + string);
        //throw new RuntimeException(getCellAddress() + " error: " + string);
    }

    private String getCellAddress() {
        return Start.cellAddress(this.formulaRow, this.formulaColumn, this.sheetName);
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
