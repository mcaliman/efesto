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


    public Parser(String filename) throws IOException, InvalidFormatException {
        super(new File(filename));
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

    @Override
    void err(String string, int row, int column) {
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
        //start.setType(internalFormulaResultTypeClass);
    }

    @Override
    protected void parseFormulaInit() {
        stack.empty();
    }

    @Override
    protected Start parseFormulaPost() {
        Start start = null;
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
    protected void _UDF(String arguments) {
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
            Reference ref = as_reference(args);
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

    private Reference as_reference(Start args) {
        if (args instanceof RangeReference) return (RangeReference) args;
        else if (args instanceof ReferenceItem) return (ReferenceItem) args;
        else if (args instanceof PrefixReferenceItem) return (PrefixReferenceItem) args;
        else return null;
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
        FLOAT term = new FLOAT(value);
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void _INT(Integer value) {
        INT term = new INT(value);
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void _TEXT(String text) {
        TEXT term = new TEXT(text);
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void _BOOL(Boolean value) {
        BOOL term = new BOOL(value);
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void _ERROR(String text) {
        ERROR term = new ERROR(text);
        setOwnProperty(term);
        err(term.toString(), formulaRow, formulaColumn);
        graph.addNode(term);
        stack.push(term);
    }


    @Override
    protected void _Plus() {
        Start expr = stack.pop();
        Plus formula = new Plus((Formula) expr);
        formula.setSheetName(currentSheetName);
        formula.setSheetIndex(currentSheetIndex);
        graph.addNode(formula);
        stack.push(formula);
    }

    @Override
    protected void _Minus() {
        Start expr = stack.pop();
        Minus formula = new Minus((Formula) expr);
        setOwnProperty(formula);
        graph.addNode(formula);
        stack.push(formula);
    }


    @Override
    protected void _Eq() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Eq op = new Eq((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Lt() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Lt op = new Lt((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Gt() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Gt op = new Gt((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Leq() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Leq op = new Leq((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _GtEq() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        GtEq op = new GtEq((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Neq() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Neq op = new Neq((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Concat() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Concat op = new Concat((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Add() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Add op = new Add((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Sub() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Sub op = new Sub((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Mult() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Mult op = new Mult((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Divide() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Divide op = new Divide((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Power() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Power op = new Power((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }


    @Override
    protected void _Intersection() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Intersection op = new Intersection((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _Union() {
        Start rExpr = stack.pop();
        Start lExpr = stack.pop();
        Union op = new Union((Formula) lExpr, (Formula) rExpr);
        setOwnProperty(op);
        graph.add(op);
        stack.push(op);
    }

    @Override
    protected void _PercentFormula() {
        PercentFormula formula = new PercentFormula((Formula) stack.pop());
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
    protected void _CELL(int ri, int ci, boolean rowNotNull, Object value, String comment) {
        CELL ref = new CELL(ri, ci);
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

    private void builtInFunction(int arity, String name) {
        try {
            if (arity == 0) {
                EXCEL_FUNCTION builtinFunction = builtinFunction(name);
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

    private EXCEL_FUNCTION builtinFunction(String name) throws UnsupportedBuiltinException {
        BuiltinFactory factory = new BuiltinFactory();
        factory.create(0, name);
        return (EXCEL_FUNCTION) factory.getBuiltInFunction();
    }

    public StartList getList() {
        return ordered;
    }

// BEGIN

    //Used
    //Sheet2!A1:B1 (Sheet + AREA/RANGE)
    @Override
    protected void parseArea3D(int firstRow, int firstColumn, int lastRow, int lastColumn, List<Object> list, String sheetName, int sheetIndex, String area) {
        SHEET tSHEET = new SHEET(sheetName);
        PrefixReferenceItem ref = new PrefixReferenceItem(tSHEET, area);
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
            FILE tFILE = new FILE(extWorkbookNumber, sheet);
            PrefixReferenceItem ref = new PrefixReferenceItem(tFILE, cellref);
            setOwnProperty(ref);
            graph.addNode(ref);
            stack.push(ref);
        } else {
            SHEET tSHEET = new SHEET(sheet);
            PrefixReferenceItem ref = new PrefixReferenceItem(tSHEET, cellref);
            setOwnProperty(ref);
            graph.addNode(ref);
            stack.push(ref);
        }

    }


    //Used

    @Override
    protected void _RangeReference(List<Object> list, int firstRow, int firstColumn, int lastRow, int lastColumn) {
        RangeReference ref = rangeReference(firstRow, firstColumn, lastRow, lastColumn);
        setOwnProperty(ref);
        //is area not a cell with ref to area
        ref.setAsArea();
        ref.add(list);
        graph.addNode(ref);
        stack.push(ref);
    }

    private RangeReference rangeReference(int firstRow, int firstColumn, int lastRow, int lastColumn) {
        CELL firstCell = new CELL(firstRow, firstColumn);
        CELL lastCell = new CELL(lastRow, lastColumn);
        return new RangeReference(firstCell, lastCell);
    }

// END
}
