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
        err("Missing ExcelFunction Arguments for cell: " + Start.cellAddress(row, column, currentSheetName), row, column);
    }

    @Override
    protected void doesFormulaReferToDeletedCell(int row, int column) {
        var address = currentSheetName + "!" + Start.cellAddress(row, column);
        err(address + " does formula refer to deleted cell", row, column);
    }

    @Override
    void err(String string, int row, int column) {
        super.err(string, row, column);
        if (errors) {
            var address = currentSheetName + "!" + Start.cellAddress(row, column);
            System.err.println(address + " ERROR: " + string);
        }
    }

    @Override
    public void parse() {
        super.parse();
        verbose("** topological sorting beginning...");
        sort();
    }


    private void sort() {
        if (unordered.singleton()) {
            ordered = new StartList();
            ordered.add(unordered.get(0));
            return;
        }
        ordered = graph.topologicalSort();
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

    // TERMINAL AND NON TERMINAL BEGIN

    /**
     * ConstantArray
     *
     * @param array
     */
    @Override
    protected void ConstantArray(Object[][] array) {
        var term = new ConstantArray(array);
        setOwnProperty(term);
        stack.push(term);
    }

    /**
     * Used
     *
     * @param arguments
     */
    @Override
    protected void UDF(String arguments) {
        var term = new UDF(arguments);
        setOwnProperty(term);
        unordered.add(term);
        stack.push(term);
    }

    /**
     * Used
     *
     * @param firstRow
     * @param firstColumn
     * @param lastRow
     * @param lastColumn
     * @param cells
     * @param name
     * @param sheetName
     */
    @Override
    protected void namedRange(int firstRow, int firstColumn, int lastRow, int lastColumn, List<Object> cells, String name, String sheetName) {
        var term = new NamedRange(name);
        term.setSheetIndex(currentSheetIndex);
        term.setSheetName(sheetName);
        term.setFirstRow(firstRow);
        term.setFirstColumn(firstColumn);
        term.setLastRow(lastRow);
        term.setLastColumn(lastColumn);
        term.setAsArea();
        term.add(cells);
        stack.push(term);
    }

    /**
     * Used
     */
    @Override
    protected void ParenthesisFormula() {
        var formula = (Formula) stack.pop();
        var parFormula = new ParenthesisFormula(formula);
        setOwnProperty(parFormula);
        stack.push(parFormula);
    }

    /**
     * Used
     *
     * @param value
     */
    @Override
    protected void FLOAT(Double value) {
        var term = new FLOAT(value);
        graph.addNode(term);
        stack.push(term);
    }

    /**
     * Used
     *
     * @param value
     */
    @Override
    protected void INT(Integer value) {
        var term = new INT(value);
        graph.addNode(term);
        stack.push(term);
    }

    /**
     * Used
     *
     * @param value
     */
    @Override
    protected void BOOL(Boolean value) {
        var term = new BOOL(value);
        graph.addNode(term);
        stack.push(term);
    }

    /**
     * Used
     *
     * @param text
     */
    @Override
    protected void TEXT(String text) {
        var term = new TEXT(text);
        graph.addNode(term);
        stack.push(term);
    }

    /**
     * Used
     *
     * @param text
     */
    @Override
    protected void ERROR(String text) {
        var term = new ERROR(text);
        setOwnProperty(term);
        err(term.toString(), formulaRow, formulaColumn);
        graph.addNode(term);
        stack.push(term);
    }

    /**
     * +
     */
    @Override
    protected void Plus() {
        var formula = (Formula) stack.pop();
        var term = new Plus(formula);
        term.setSheetName(currentSheetName);
        term.setSheetIndex(currentSheetIndex);
        graph.addNode(term);
        stack.push(term);
    }

    /**
     * -
     */
    @Override
    protected void Minus() {
        var formula = (Formula) stack.pop();
        var term = new Minus(formula);
        setOwnProperty(term);
        graph.addNode(term);
        stack.push(term);
    }

    /**
     * F=F
     */
    @Override
    protected void eq() {
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
    @Override
    protected void lt() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var lt = new Lt(lFormula, rFormula);
        setOwnProperty(lt);
        graph.add(lt);
        stack.push(lt);
    }

    /**
     * F>F
     */
    @Override
    protected void gt() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var gt = new Gt(lFormula, rFormula);
        setOwnProperty(gt);
        graph.add(gt);
        stack.push(gt);
    }

    /**
     * F<=F
     */
    @Override
    protected void leq() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var leq = new Leq(lFormula, rFormula);
        setOwnProperty(leq);
        graph.add(leq);
        stack.push(leq);
    }

    /**
     * F>=F
     */
    @Override
    protected void gteq() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        GtEq gteq = new GtEq(lFormula, rFormula);
        setOwnProperty(gteq);
        graph.add(gteq);
        stack.push(gteq);
    }

    /**
     * F<>F
     */
    @Override
    protected void neq() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var neq = new Neq(lFormula, rFormula);
        setOwnProperty(neq);
        graph.add(neq);
        stack.push(neq);
    }

    /**
     * F&F
     */
    @Override
    protected void concat() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var concat = new Concat(lFormula, rFormula);
        setOwnProperty(concat);
        graph.add(concat);
        stack.push(concat);
    }

    /**
     * F+F
     */
    @Override
    protected void add() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var add = new Add(lFormula, rFormula);
        setOwnProperty(add);
        graph.add(add);
        stack.push(add);
    }

    /**
     * F-F
     */
    @Override
    protected void sub() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var sub = new Sub(lFormula, rFormula);
        setOwnProperty(sub);
        graph.add(sub);
        stack.push(sub);
    }

    /**
     * F*F
     */
    @Override
    protected void mult() {
        if (stack.empty()) return;
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var mult = new Mult(lFormula, rFormula);
        setOwnProperty(mult);
        graph.add(mult);
        stack.push(mult);
    }

    /**
     * F/F
     */
    @Override
    protected void div() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var div = new Divide(lFormula, rFormula);
        setOwnProperty(div);
        graph.add(div);
        stack.push(div);
    }

    /**
     * F^F
     */
    @Override
    protected void power() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var power = new Power(lFormula, rFormula);
        setOwnProperty(power);
        graph.add(power);
        stack.push(power);
    }


    /**
     * F F
     */
    @Override
    protected void intersection() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var intersection = new Intersection(lFormula, rFormula);
        setOwnProperty(intersection);
        graph.add(intersection);
        stack.push(intersection);
    }

    /**
     * F,F
     */
    @Override
    protected void union() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var union = new Union(lFormula, rFormula);
        setOwnProperty(union);
        graph.add(union);
        stack.push(union);
    }

    /**
     * F%
     */
    @Override
    protected void percentFormula() {
        var formula = (Formula) stack.pop();
        var percentFormula = new PercentFormula(formula);
        setOwnProperty(percentFormula);
        graph.addNode(percentFormula);
        stack.push(percentFormula);
    }

    /**
     * #REF
     *
     * @param text
     */
    @Override
    protected void ERROR_REF(String text) {
        var error = new ERROR_REF();
        setOwnProperty(error);
        stack.push(error);
        err(text, formulaRow, formulaColumn);
    }

    /**
     * CELLREF
     *
     * @param ri
     * @param ci
     * @param rowNotNull
     * @param value
     * @param comment
     */
    @Override
    protected void CELL_REFERENCE(int ri, int ci, boolean rowNotNull, Object value, String comment) {
        var term = new CELL_REFERENCE(ri, ci);
        term.setComment(comment);
        setOwnProperty(term);
        if (rowNotNull) {
            term.setValue(value);
            this.unordered.add(term);
        }
        stack.push(term);
    }

    /**
     * Used
     * Sheet2!A1:B1 (Sheet + AREA/RANGE)
     *
     * @param firstRow
     * @param firstColumn
     * @param lastRow
     * @param lastColumn
     * @param list
     * @param area
     */
    @Override
    protected void parseArea3D(RANGE tRANGE, int firstRow, int firstColumn, int lastRow, int lastColumn, List<Object> list, SHEET tSHEET, String area) {
        var term = new PrefixReferenceItem(tSHEET, area, tRANGE);
        term.setSheetIndex(tSHEET.getIndex());
        term.setSheetName(tSHEET.getName());
        term.setAsArea();
        term.add(list);
        term.setFirstRow(firstRow);
        term.setFirstColumn(firstColumn);
        term.setLastRow(lastRow);
        term.setLastColumn(lastColumn);
        unordered.add(term);
        stack.push(term);
    }

    /**
     * Used
     * Sheet2!A1 (Sheet + CELL_REFERENCE)
     * External references: External references are normally in the form [File]Sheet!Cell
     *
     * @param cellref
     */
    @Override
    protected void parseReference(FILE tFILE, String cellref) {
        var term = new PrefixReferenceItem(tFILE, cellref, null);
        setOwnProperty(term);
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void parseReference(SHEET tSHEET, String cellref) {
        var term = new PrefixReferenceItem(tSHEET, cellref, null);
        setOwnProperty(term);
        graph.addNode(term);
        stack.push(term);
    }

    /**
     * Used
     */
    @Override
    protected void rangeReference(RANGE tRANGE) {
        var term = new RangeReference(tRANGE.getFirst(), tRANGE.getLast());
        setOwnProperty(term);
        term.setAsArea();//is area not a cell with ref to area
        term.add(tRANGE.values()/* list*/);
        graph.addNode(term);
        stack.push(term);
    }

    /**
     * SUM(Arguments)
     */
    @Override
    protected void sum() {
        var args = stack.pop();
        if (args instanceof Reference || args instanceof OFFSET) {
            args.setSheetIndex(currentSheetIndex);
            args.setSheetName(currentSheetName);
            args.setAsArea();
            unordered.add(args);
        } else {
            err("Not RangeReference " + args.getClass().getSimpleName() + " " + args.toString(), formulaRow, formulaColumn);
        }
        var term = new SUM((Formula) args);
        setOwnProperty(term);
        unordered.add(term);
        graph.add(term);
        stack.push(term);
    }


    /**
     * @param name
     * @param arity
     * @param externalFunction
     */
    @Override
    protected void parseFunc(String name, int arity, boolean externalFunction) {
        try {
            if (arity == 0) builtinFunction(name);
            else builtInFunction(arity, name);
        } catch (UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e, formulaRow, formulaColumn);
        }
    }

    // TERMINAL AND NON TERMINAL END

    /**
     * @param arity
     * @param name
     */
    private void builtInFunction(int arity, String name) throws UnsupportedBuiltinException {
        var factory = new BuiltinFactory();
        factory.create(arity, name);
        var builtinFunction = (EXCEL_FUNCTION) factory.getBuiltInFunction();
        Start[] args = factory.getArgs();
        for (int i = arity - 1; i >= 0; i--) if (!stack.empty()) args[i] = stack.pop();

        setOwnProperty(builtinFunction);
        graph.addNode(builtinFunction);
        for (Start arg : args) {
            if (arg instanceof RangeReference || arg instanceof CELL_REFERENCE || arg instanceof PrefixReferenceItem || arg instanceof ReferenceItem) {
                if (unordered.add(arg)) {
                    graph.addNode(arg);
                    graph.addEdge(arg, builtinFunction);
                }
            }
        }
        stack.push(builtinFunction);
    }

    private void builtinFunction(String name) throws UnsupportedBuiltinException {
        BuiltinFactory factory = new BuiltinFactory();
        factory.create(0, name);
        var builtinFunction = (EXCEL_FUNCTION) factory.getBuiltInFunction();
        stack.push(builtinFunction);
    }

    public StartList getList() {
        return ordered;
    }
}
