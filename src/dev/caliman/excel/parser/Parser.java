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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

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


/**
 * @author Massimo Caliman
 */
public final class Parser extends AbstractParser {


    private boolean verbose = false;
    private Helper helper;
    private List<Cell> ext;


    private StartList unordered;
    private StartList ordered;
    private StartGraph graph;
    private Stack<Start> stack;

    public Parser(String filename) throws IOException, InvalidFormatException {
        super(filename);
        this.ext = new ArrayList<>();
        this.helper = new Helper(this.workbook);
        this.unordered = new StartList();
        this.ordered = new StartList();
        this.graph = new StartGraph();
        this.stack = new Stack<>();
    }


    protected void parse(Cell cell) {
        if(isFormula(cell)) {
            parseFormula(cell);
        } else if(this.ext.contains(cell)) {
            verbose("Recover loosed cell!");
            Object value = this.parseCellValue(cell);
            CELL elem = new CELL(cell.getRowIndex(), cell.getColumnIndex());
            elem.setValue(value);
            elem.setSHEET(new SHEET(getSheetName(cell), getSheetIndex(cell)));
            parseCELLlinked(elem);
            this.ext.remove(cell);
        } else if(!this.ext.contains(cell) && !empty(cell)) {
            //Non è formula non è nelle celle utili collezionate
            out.println("Cella di interesse? " + cell.toString());
        }
    }


    protected void parseFormula(Cell cell) {
        super.parseFormula(cell);
        if(this.formulaPtgs == null) {
            err("ptgs empty or null for address " + this.formulaAddress);
            parseUDF(this.formulaPlainText);
            return;
        }
        Start start = parse(this.formulaPtgs);
        if(start != null) {
            start.setSingleSheet(this.singleSheet);
            parseFormula(start);
        }
    }


    private Start parse(Ptg[] ptgs) {
        stack.empty();
        if(Ptg.doesFormulaReferToDeletedCell(ptgs)) doesFormulaReferToDeletedCell();
        for(Ptg ptg : ptgs) parse(ptg);
        Start start = null;
        if(!stack.empty()) start = stack.pop();
        return start;
    }

    private void parseUDF(String arguments) {
        var elem = new UDF(arguments);
        elem.setColumn(this.column);
        elem.setRow(this.row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        unordered.add(elem);
        stack.push(elem);
    }

    private void parse(Ptg p) {
        verbose("parse: " + p.getClass().getSimpleName());
        try(Stream<WhatIf> stream = Stream.of(
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
        } catch(Exception e) {
            //err.println("parse: " + p.getClass().getSimpleName() + " " + this.cSHEET.getName() + "row:" + row + "column:" + column + e.getMessage());
            e.printStackTrace();
        }
    }


    private void parseArea3DPxg(Area3DPxg t) {
        // Area3DPxg is XSSF Area 3D Reference (Sheet + Area) Defined an area in an
        // external or different sheet.
        // This is XSSF only, as it stores the sheet / workbook references in String
        // form. The HSSF equivalent using indexes is Area3DPtg
        String name = t.getSheetName();
        int index = getSheetIndex(name);
        SHEET tSHEET = new SHEET(name, index);
        String area = helper.getArea(t);
        parseArea3D(helper.getRANGE(name, t), tSHEET, area);
    }

    /**
     * Sheet2!A1:B1 (Sheet + AREA/RANGE)
     */
    private void parseArea3D(RANGE tRANGE, SHEET tSHEET, String area) {
        var elem = new PrefixReferenceItem(tSHEET, area, tRANGE);
        elem.setSHEET(tSHEET);
        unordered.add(elem);
        stack.push(elem);
    }

    private void parseRef3DPxg(Ref3DPxg t) {
        //Title: XSSF 3D Reference
        //Description: Defines a cell in an external or different sheet.
        //REFERENCE:
        //This is XSSF only, as it stores the sheet / workbook references in String form.
        //The HSSF equivalent using indexes is Ref3DPtg
        int extWorkbookNumber = t.getExternalWorkbookNumber();
        String sheetName = t.getSheetName();
        int sheetIndex = getSheetIndex(sheetName);
        SHEET tSHEET = new SHEET(sheetName, sheetIndex);
        FILE tFILE = new FILE(extWorkbookNumber, tSHEET);
        String cellref = helper.getCellRef(t);
        if(this.getSheetIndex() != sheetIndex) {
            Sheet extSheet = this.workbook.getSheet(sheetName);
            if(extSheet != null) {
                CellReference cr = new CellReference(cellref);
                Row row = extSheet.getRow(cr.getRow());
                Cell cell = row.getCell(cr.getCol());
                this.ext.add(cell);
                verbose("Loosing!!! reference[ext] " + tSHEET.toString() + "" + cellref);
            }
        }
        if(extWorkbookNumber > 0) parseReference(tFILE, cellref);
        else parseReference(tSHEET, cellref);
    }


    private void parseReference(SHEET tSHEET, String cellref) {
        var elem = new PrefixReferenceItem(tSHEET, cellref, null);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.addNode(elem);
        stack.push(elem);
    }

    /**
     * RangeReference
     */
    private void parseAreaPtg(AreaPtg t) {
        RANGE tRANGE = helper.getRANGE(sheet, t);
        var elem = new RangeReference(tRANGE.getFirst(), tRANGE.getLast());
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);

        elem.setAsArea();//is area not a cell with ref to area
        elem.add(tRANGE.values());
        graph.addNode(elem);
        stack.push(elem);

    }

    private void parseNamedRange(NamePtg t) {
        RangeInternal range = null;
        Ptg[] ptgs = getName(t);
        String name = getNameText(t);
        int sheetIndex = 0;
        for(Ptg ptg : ptgs) {
            if(ptg != null) {
                if(ptg instanceof Area3DPxg) {
                    Area3DPxg area3DPxg = (Area3DPxg) ptg;
                    range = new RangeInternal(workbook, area3DPxg.getSheetName(), area3DPxg);
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

    private void parseRefPtg(RefPtg t) {
        Row row_ = this.sheet.getRow(t.getRow());
        Object value = null;
        if(row_ != null) {
            Cell c = row_.getCell(t.getColumn());
            value = Helper.valueOf(c);
        }
        CELL elem = new CELL(t.getRow(), t.getColumn());
        elem.setValue(value);
        elem.setColumn(this.column);
        elem.setRow(this.row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        this.unordered.add(elem);
        stack.push(elem);
    }

    /**
     * ConstantArray
     */
    private void parseConstantArray(ArrayPtg t) {
        Object[][] array = t.getTokenArrayValues();
        var elem = new ConstantArray(array);
        elem.setColumn(this.column);
        elem.setRow(this.row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        this.stack.push(elem);
    }

    private void parseAttrPtg(AttrPtg t) {
        if(t.isSum()) parseSum();
    }

    /**
     * SUM(Arguments)
     */
    private void parseSum() {
        var args = stack.pop();
        if(args instanceof Reference || args instanceof OFFSET) {
            args.setSheetIndex(this.getSheetIndex());
            args.setSheetName(this.getSheetName());
            args.setAsArea();
            unordered.add(args);
        } else {
            err("Not RangeReference " + args.getClass().getSimpleName() + " " + args.toString());
        }
        var elem = new SUM((Formula) args);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        unordered.add(elem);
        graph.add(elem);
        stack.push(elem);
    }

    private void parseFuncVarPtg(FuncVarPtg t) {
        int arity = t.getNumberOfOperands();
        String name = t.getName();
        if(arity == 0) parseFunc(name);
        else parseFunc(name, arity);
    }

    private void parseFuncPtg(FuncPtg t) {
        int arity = t.getNumberOfOperands();
        String name = t.getName();
        if(arity == 0) parseFunc(name);
        else parseFunc(name, arity);
    }

    private void parseFunc(String name, int arity) {
        try {
            builtInFunction(arity, name);
        } catch(UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e);
        }
    }

    //@todo impl. DATE
    private void parseERROR(ErrPtg t) {
        String text;
        if(t == NULL_INTERSECTION) text = "#NULL!";
        else if(t == DIV_ZERO) text = "#DIV/0!";
        else if(t == VALUE_INVALID) text = "#VALUE!";
        else if(t == REF_INVALID) text = "#REF!";
        else if(t == NAME_INVALID) text = "#NAME?";
        else if(t == NUM_ERROR) text = "#NUM!";
        else if(t == N_A) text = "#N/A";
        else text = "FIXME!";


        // ERROR
        var elem = new ERROR(text);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
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
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        stack.push(elem);
        err("");
    }


    public void sort() {
        if(this.unordered.singleton()) {
            this.ordered = new StartList();
            this.ordered.add(this.unordered.get(0));
            return;
        }
        this.ordered = this.graph.topologicalSort();
    }

    private void parseFormula(Start elem) {
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        unordered.add(elem);
    }


    private void parseCELLlinked(CELL elem) {
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
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
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.addNode(elem);
        stack.push(elem);
    }


    private void parseFunc(String name) {
        try {
            builtinFunction(name);
        } catch(UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e);
        }
    }


    private void builtInFunction(int arity, String name) throws UnsupportedBuiltinException {
        var factory = new BuiltinFactory();
        factory.create(arity, name);
        var builtinFunction = (EXCEL_FUNCTION) factory.getBuiltInFunction();
        Start[] args = factory.getArgs();
        for(int i = arity - 1; i >= 0; i--) if(!stack.empty()) args[i] = stack.pop();

        builtinFunction.setColumn(column);
        builtinFunction.setRow(row);
        builtinFunction.setSheetIndex(this.getSheetIndex());
        builtinFunction.setSheetName(this.getSheetName());
        builtinFunction.setSingleSheet(this.singleSheet);

        graph.addNode(builtinFunction);
        for(Start arg : args) {
            if(arg instanceof RangeReference /*|| arg instanceof CELL*/ || arg instanceof PrefixReferenceItem || arg instanceof ReferenceItem) {
                if(unordered.add(arg)) {
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


    private void err(String string) {
        err.println(getCellAddress() + " error: " + string);
        //throw new RuntimeException(getCellAddress() + " error: " + string);
    }


    //<editor-fold desc="GRAMMAR ELEMENTS">

    /**
     * (F)
     */
    private void parseParenthesisFormula() {
        var formula = (Formula) stack.pop();
        var elem = new ParenthesisFormula(formula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        stack.push(elem);
    }

    /**
     * F=F
     */
    private void parseEq() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Eq(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F<F
     */
    private void parseLt() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Lt(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F>F
     */
    private void parseGt() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Gt(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F<=F
     */
    private void parseLeq() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Leq(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F>=F
     */
    private void parseGteq() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new GtEq(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F<>F
     */
    private void parseNeq() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Neq(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F&F
     */
    private void parseConcat() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Concat(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F+F
     */
    private void parseAdd() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Add(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F-F
     */
    private void parseSub() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Sub(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F*F
     */
    private void parseMult() {
        if(stack.empty()) return;
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Mult(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F/F
     */
    private void parseDiv() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Divide(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F^F
     */
    private void parsePower() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Power(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }

    /**
     * F%
     */
    private void percentFormula() {
        var formula = (Formula) stack.pop();
        var elem = new PercentFormula(formula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.addNode(elem);
        stack.push(elem);
    }

    /**
     * + F
     */
    private void parsePlus() {
        var formula = (Formula) stack.pop();
        var elem = new Plus(formula);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        graph.addNode(elem);
        stack.push(elem);
    }

    /**
     * - F
     */
    private void parseMinus() {
        var formula = (Formula) stack.pop();
        var elem = new Minus(formula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.addNode(elem);
        stack.push(elem);
    }

    /**
     * Intersection
     * F F
     */
    private void parseIntersection() {
        var rFormula = (Formula) this.stack.pop();
        var lFormula = (Formula) this.stack.pop();
        var elem = new Intersection(lFormula, rFormula);
        elem.setColumn(this.column);
        elem.setRow(this.row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        this.graph.add(elem);
        this.stack.push(elem);
    }

    /**
     * Union
     * F,F
     */
    private void parseUnion() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var elem = new Union(lFormula, rFormula);
        elem.setColumn(column);
        elem.setRow(row);
        elem.setSheetIndex(this.getSheetIndex());
        elem.setSheetName(this.getSheetName());
        elem.setSingleSheet(this.singleSheet);
        graph.add(elem);
        stack.push(elem);
    }
    //</editor-fold>

    //<editor-fold desc="Utilities">


    public void setVerbose(boolean verbose) {
        this.verbose = verbose;
    }

    public int getCounterFormulas() {
        return noOfFormulas;
    }


    private void verbose(String text) {
        if(this.verbose) out.println(text);
    }


    private int getSheetIndex(String xlsxSheetName) {
        return helper.getSheetIndex(xlsxSheetName);
    }

    private int getSheetIndex(Cell xlsxCell) {
        return helper.getSheetIndex(xlsxCell.getSheet().getSheetName());
    }


    //</editor-fold>


    // INNERS CLASS

    protected class WhatIf {

        final Ptg ptg;
        final Predicate<Ptg> predicate;
        final Consumer<Ptg> consumer;

        WhatIf(Ptg ptg, Predicate<Ptg> predicate, Consumer<Ptg> consumer) {
            this.ptg = ptg;
            this.predicate = predicate;
            this.consumer = consumer;
        }
    }

    protected class RangeInternal {

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

            for(Cell cell : cells)
                if(cell != null) {
                    tRANGE.add(Helper.valueOf(cell));
                }
        }


        String getSheetName() {
            return sheetName;
        }


        RANGE getRANGE() {
            return tRANGE;
        }

    }

}
