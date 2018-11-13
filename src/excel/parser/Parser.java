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

package excel.parser;

import excel.grammar.Formula;
import excel.grammar.Grammar;
import excel.grammar.Metadata;
import excel.grammar.Start;
import excel.grammar.formula.ConstantArray;
import excel.grammar.formula.ParenthesisFormula;
import excel.grammar.formula.Reference;
import excel.grammar.formula.constant.*;
import excel.grammar.formula.functioncall.EXCEL_FUNCTION;
import excel.grammar.formula.functioncall.PercentFormula;
import excel.grammar.formula.functioncall.binary.*;
import excel.grammar.formula.functioncall.builtin.SUM;
import excel.grammar.formula.functioncall.unary.Minus;
import excel.grammar.formula.functioncall.unary.Plus;
import excel.grammar.formula.reference.*;
import excel.grammar.formula.reference.referencefunction.OFFSET;
import excel.graph.StartGraph;
import excel.parser.internal.AbstractParser;
import excel.parser.internal.HelperInternal;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.Stack;

/**
 * @author Massimo Caliman
 */
public final class Parser extends AbstractParser {

    private StartList unordered;
    private StartList ordered;
    private StartGraph graph;
    private Stack<Start> stack;
    private Grammar grammar;

    public Parser(String filename) throws IOException, InvalidFormatException {
        super(new File(filename));
        grammar = new Grammar();
        unordered = new StartList();
        ordered = new StartList();
        graph = new StartGraph();
        stack = new Stack<>();
    }

    @Override
    protected void parseMissingArguments(int row, int column) {
        err("Missing ExcelFunction Arguments for cell: " + HelperInternal.cellAddress(row, column, currentSheetName), row, column);
    }

    @Override
    protected void doesFormulaReferToDeletedCell(int row, int column) {
        String address = currentSheetName + "!" + HelperInternal.cellAddress(row, column);
        err(address + " does formula refer to deleted cell", row, column);
    }

    /*private void unmanaged(String text) {
        System.err.println("unmanaged!" + text);
        throw new RuntimeException("unmanaged!" + text);
    }*/

    @Override
    protected void err(String string, int row, int column) {
        super.err(string, row, column);
        if (errors) {
            String address = currentSheetName + "!" + HelperInternal.cellAddress(row, column);
            System.err.println(address + " ERROR: " + string);
        }
    }

    @Override
    public void parse() {
        super.parse();
        verbose("** topological sorting beginning...");
        sort();
        metadata();
    }


    private void sort() {
        if (unordered.singleton()) {
            ordered = new StartList();
            ordered.add(unordered.get(0));
            return;
        }
        ordered = graph.topologicalSort();
    }

    private void metadata() {
        if (metadata) {
            ordered.add(0, new Metadata("filename", fileName));
            //this.ordered.add(1, new Metadata("creator", this.creator));
            //this.ordered.add(2, new Metadata("description", this.description));
            //this.ordered.add(3, new Metadata("keywords", this.keywords));
            //this.ordered.add(4, new Metadata("title", this.title));
            //this.ordered.add(5, new Metadata("subject", this.subject));
            //this.ordered.add(6, new Metadata("category", this.category));
            //this.ordered.add(7, new Metadata("author", this.author));
            //this.ordered.add(8, new Metadata("company", this.company));
            //this.ordered.add(9, new Metadata("template", this.template));
            //this.ordered.add(10, new Metadata("template", this.template));
            //this.ordered.add(11, new Metadata("manager", this.manager));
            //this.ordered.add(12, new Metadata("", ""));
        }
    }

    @Override
    public void parseFormula(Start obj) {
        if (Objects.isNull(obj)) return;
        setOwnProperty(obj);
        unordered.add(obj);
    }

    private void setOwnProperty(Start start) {
        start.setColumn(formulaColumn);
        start.setRow(formulaRow);
        start.setSheetIndex(currentSheetIndex);
        start.setSheetName(currentSheetName);
        start.setType(internalFormulaResultTypeClass);
    }

    @Override
    protected void parseFormulaInit() {
        stack.empty();
    }

    @Override
    protected Start parseFormulaPost(Start start, int row, int column) {
        if (!stack.empty()) start = stack.pop();
        return start;
    }

    //Used
    @Override
    protected void _ConstantArray(Object[][] array) {
        ConstantArray term = new ConstantArray(array);
        setOwnProperty(term);
        stack.push(term);
    }


    //Used
    @Override
    protected void _UDF(String arguments, int formulaRow, int formulaColumn) {
        UDF udf = new UDF(arguments);
        setOwnProperty(udf);
        unordered.add(udf);
        stack.push(udf);
    }

    //Used
    @Override
    protected void _SUM() {
        Start args = stack.pop();
        if (args instanceof Reference) {
            Reference ref = grammar.as_reference(args);
            ref.setSheetIndex(currentSheetIndex);
            ref.setSheetName(currentSheetName);
            ref.setAsArea();
            unordered.add(ref);
        } else if (args instanceof OFFSET) {
            OFFSET ref = (OFFSET) args;
            ref.setSheetIndex(currentSheetIndex);
            ref.setSheetName(currentSheetName);
            ref.setAsArea();
            unordered.add(ref);
        } else
            err("Not RangeReference " + args.getClass().getSimpleName() + " " + args.toString(), formulaRow, formulaColumn);
        SUM sum = new SUM((Formula) args);
        setOwnProperty(sum);
        unordered.add(sum);
        graph.add(sum);
        stack.push(sum);
    }

    //Used
    @Override
    protected void _NamedRange(int firstRow, int firstColumn, int lastRow, int lastColumn, List<Object> cells, String name, String sheetName) {
        NamedRange ref = new NamedRange(name);
        ref.setSheetIndex(currentSheetIndex);
        ref.setSheetName(sheetName);
        ref.setFirstRow(firstRow);
        ref.setFirstColumn(firstColumn);
        ref.setLastRow(lastRow);
        ref.setLastColumn(lastColumn);
        ref.setAsArea();
        ref.add(cells);
        stack.push(ref);
    }

    //Used
    @Override
    protected void _ParenthesisFormula() {
        Start obj = stack.pop();
        ParenthesisFormula formula = new ParenthesisFormula((Formula) obj);
        setOwnProperty(formula);
        stack.push(formula);
    }

    @Override
    protected void _FLOAT(Double value) {
        FLOAT term = grammar.number(value);
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void _INT(Integer value) {
        INT term = grammar.number(value);
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void _TEXT(String text) {
        TEXT term = grammar.text(text);
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void _BOOL(Boolean value) {
        BOOL term = grammar.bool(value);
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void _ERROR(String text) {
        ERROR term = grammar.error(text);
        setOwnProperty(term);
        err(term.toString(), formulaRow, formulaColumn);
        graph.addNode(term);
        stack.push(term);
    }


    @Override
    protected void _Plus() {
        Start expr = stack.pop();
        Plus formula = grammar.plus(expr);
        formula.setSheetName(currentSheetName);
        formula.setSheetIndex(currentSheetIndex);
        graph.addNode(formula);
        stack.push(formula);
    }

    @Override
    protected void _Minus() {
        Start expr = stack.pop();
        Minus formula = grammar.minus(expr);
        setOwnProperty(formula);
        graph.addNode(formula);
        stack.push(formula);
    }


    @Override
    protected void _Eq() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Eq op = grammar.eq(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Lt() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Lt op = grammar.lt(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Gt() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Gt op = grammar.gt(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Leq() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Leq op = grammar.leq(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _GtEq() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        GtEq op = grammar.gteq(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Neq() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Neq op = grammar.neq(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Concat() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Concat op = grammar.concat(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Add() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Add op = grammar.add(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Sub() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Sub op = grammar.subtrac(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Mult() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Mult op = grammar.multiply(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Divide() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Divide op = grammar.divide(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Power() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Power op = grammar.power(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }


    @Override
    protected void _Intersection() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Intersection op = grammar.intersection(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Union() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Union op = grammar.union(lExpr, rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _PercentFormula() {
        PercentFormula formula = grammar.percentFormula(stack.pop());
        setOwnProperty(formula);
        graph.addNode(formula);
        stack.push(formula);
    }

    @Override
    protected void _ERROR_REF(String text) {
        ERROR_REF ref = new ERROR_REF();
        setOwnProperty(ref);
        stack.push(ref);
        err(text, formulaRow, formulaColumn);
    }

    @Override
    protected void _CELL(int ri, int ci, boolean rowRelative, boolean colRelative, boolean rowNotNull, Object value, String comment) {
        CELL ref = grammar.cell(ri, ci, rowRelative, colRelative);
        ref.setComment(comment);
        setOwnProperty(ref);
        if (rowNotNull) {
            ref.setValue(value);
            this.unordered.add(ref);
        }
        stack.push(ref);
    }

    //Used
    @Override
    protected void parseFunc(String name, int arity) {
        builtInFunction(arity, name);
    }

    //Used (unique with parseFunc)
    @Override
    protected void parseFuncVar(String name, int arity) {
        builtInFunction(arity, name);
    }

    protected void builtInFunction(int arity, String name) {
        try {
            if (arity == 0) {
                EXCEL_FUNCTION builtinFunction = grammar.builtinFunction(name);
                stack.push(builtinFunction);
                return;
            }
        } catch (UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e, formulaRow, formulaColumn);
        }

        BuiltinFactory factory = new BuiltinFactory();
        try {
            factory.create(arity, name);
            EXCEL_FUNCTION builtinFunction = (EXCEL_FUNCTION) factory.getBuiltInFunction();
            Start[] args = factory.getArgs();
            for (int i = arity - 1; i >= 0; i--) if (!stack.empty()) args[i] = stack.pop();
            setOwnProperty(builtinFunction);
            graph.addNode(builtinFunction);
            for (Start arg : args) {
                if (arg instanceof RangeReference) {
                    if (unordered.add(arg)) {
                        graph.addNode(arg);
                        graph.addEdge(arg, builtinFunction);
                    }
                } else if (arg instanceof CELL) {
                    if (unordered.add(arg)) {
                        graph.addNode(arg);
                        graph.addEdge(arg, builtinFunction);
                    }
                } else if (arg instanceof PrefixReferenceItem) {
                    if (unordered.add(arg)) {
                        graph.addNode(arg);
                        graph.addEdge(arg, builtinFunction);
                    }

                } else if (arg instanceof ReferenceItem) {
                    if (unordered.add(arg)) {
                        graph.addNode(arg);
                        graph.addEdge(arg, builtinFunction);
                    }
                }
            }
            stack.push(builtinFunction);
        } catch (UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e, formulaRow, formulaColumn);
        }
    }

    public StartList getList() {
        return ordered;
    }

// BEGIN

    //Used
    //Sheet2!A1:B1 (Sheet + AREA/RANGE)
    @Override
    protected void parseArea3D(int firstRow, int firstColumn, int lastRow, int lastColumn, List<Object> list, String sheetName, int sheetIndex, String area) {
        PrefixReferenceItem ref = new PrefixReferenceItem(sheetName, area);
        ref.setSheetIndex(sheetIndex);
        ref.setSheetName(sheetName);
        ref.setAsArea();
        ref.add(list);
        ref.setFirstRow(firstRow);
        ref.setFirstColumn(firstColumn);
        ref.setLastRow(lastRow);
        ref.setLastColumn(lastColumn);
        unordered.add(ref);
        stack.push(ref);
    }

    //Used
    //Sheet2!A1 (Sheet + CELL)
    @Override
    protected void parseRef3D(int extWorkbookNumber, String sheet, String cellref) {
        //External references: External references are normally in the form [File]Sheet!Cell
        if (extWorkbookNumber > 0) {
            FILE tFILE = new FILE(extWorkbookNumber);
            PrefixReferenceItem ref = new PrefixReferenceItem(tFILE.toString() + sheet, cellref);
            setOwnProperty(ref);
            graph.addNode(ref);
            stack.push(ref);
        } else {
            SHEET tSHEET = new SHEET(sheet);
            PrefixReferenceItem ref = new PrefixReferenceItem(tSHEET.toString(), cellref);
            setOwnProperty(ref);
            graph.addNode(ref);
            stack.push(ref);
        }

    }


    //Used

    @Override
    protected void _RangeReference(List<Object> list, int firstRow, int firstColumn, boolean isFirstRowRelative, boolean isFirstColRelative, int lastRow, int lastColumn, boolean isLastRowRelative, boolean isLastColRelative) {
        RangeReference ref = grammar.rangeReference(firstRow, firstColumn, isFirstRowRelative, isFirstColRelative, lastRow, lastColumn, isLastRowRelative, isLastColRelative);
        setOwnProperty(ref);
        //is area not a cell with ref to area
        ref.setAsArea();
        ref.add(list);
        graph.addNode(ref);
        stack.push(ref);
    }

// END
}
