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
import excel.parser.internal.AbstractParser;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;
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
        err("Missing ExcelFunction Arguments for cell: " + Start.cellAddress(row, column, sheetName), row, column);
    }

    @Override
    protected void doesFormulaReferToDeletedCell(int row, int column) {
        err(Start.cellAddress(row, column, sheetName) + " does formula refer to deleted cell", row, column);
    }

    @Override
    protected void err(String string, int row, int column) {
        super.err(string, row, column);
        if (errors) System.err.println(Start.cellAddress(row, column, sheetName) + " parseErrorLiteral: " + string);
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
        start.setColumn(colFormula);
        start.setRow(rowFormula);
        start.setSheetIndex(sheetIndex);
        start.setSheetName(sheetName);
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
     * parseConstantArray
     *
     * @param array
     */
    @Override
    protected void parseConstantArray(Object[][] array) {
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
    protected void parseUDF(String arguments) {
        var term = new UDF(arguments);
        setOwnProperty(term);
        unordered.add(term);
        stack.push(term);
    }

    /**
     * Used
     *
     * @param name
     * @param sheetName
     */
    @Override
    protected void parseNamedRange(RANGE tRANGE, String name, String sheetName) {
        var term = new NamedRange(name, tRANGE);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(sheetName);
        stack.push(term);
    }

    /**
     * Used
     */
    @Override
    protected void parseParenthesisFormula() {
        var formula = (Formula) stack.pop();
        var parFormula = new ParenthesisFormula(formula);
        setOwnProperty(parFormula);
        stack.push(parFormula);
    }


    /**
     * F=F
     */
    @Override
    protected void parseEq() {
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
    protected void parseLt() {
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
    protected void parseGt() {
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
    protected void parseLeq() {
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
    protected void parseGteq() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var gteq = new GtEq(lFormula, rFormula);
        setOwnProperty(gteq);
        graph.add(gteq);
        stack.push(gteq);
    }

    /**
     * F<>F
     */
    @Override
    protected void parseNeq() {
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
    protected void parseConcat() {
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
    protected void parseAdd() {
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
    protected void parseSub() {
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
    protected void parseMult() {
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
    protected void parseDiv() {
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
    protected void parsePower() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var power = new Power(lFormula, rFormula);
        setOwnProperty(power);
        graph.add(power);
        stack.push(power);
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
     */
    @Override
    protected void parseReferenceErrorLiteral(ERROR_REF error) {
        setOwnProperty(error);
        stack.push(error);
        err("", rowFormula, colFormula);
    }

    /**
     * CELLREF
     */
    @Override
    protected void parseCELL_REFERENCE(CELL_REFERENCE tCELL_REFERENCE, boolean rowNotNull, Object value) {
        setOwnProperty(tCELL_REFERENCE);
        if (rowNotNull) {
            tCELL_REFERENCE.setValue(value);
            this.unordered.add(tCELL_REFERENCE);
        }
        stack.push(tCELL_REFERENCE);
    }

    /**
     * Used
     * Sheet2!A1:B1 (Sheet + AREA/RANGE)
     */
    @Override
    protected void parseArea3D(RANGE tRANGE, SHEET tSHEET, String area) {
        var term = new PrefixReferenceItem(tSHEET, area, tRANGE);
        term.setSheetIndex(tSHEET.getIndex());
        term.setSheetName(tSHEET.getName());
        unordered.add(term);
        stack.push(term);
    }

    /**
     * Used
     * Sheet2!A1 (Sheet + parseCELL_REFERENCE)
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
    protected void parseRangeReference(RANGE tRANGE) {
        var rangeReference = new RangeReference(tRANGE.getFirst(), tRANGE.getLast());
        setOwnProperty(rangeReference);
        rangeReference.setAsArea();//is area not a cell with ref to area
        rangeReference.add(tRANGE.values());
        graph.addNode(rangeReference);
        stack.push(rangeReference);
    }

    /**
     * SUM(Arguments)
     */
    @Override
    protected void parseSum() {
        var args = stack.pop();
        if (args instanceof Reference || args instanceof OFFSET) {
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


    @Override
    protected void parseFunc(String name, boolean externalFunction) {
        try {
            builtinFunction(name);
        } catch (UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e, rowFormula, colFormula);
        }
    }

    @Override
    protected void parseFunc(String name, int arity, boolean externalFunction) {
        try {
            builtInFunction(arity, name);
        } catch (UnsupportedBuiltinException e) {
            err("Unsupported Excel ExcelFunction: " + name + " " + e, rowFormula, colFormula);
        }
    }

    // TERMINAL AND NON TERMINAL END

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
        var factory = new BuiltinFactory();
        factory.create(0, name);
        var builtinFunction = (EXCEL_FUNCTION) factory.getBuiltInFunction();
        stack.push(builtinFunction);
    }

    public StartList getList() {
        return ordered;
    }


//Unary BEGIN

    /**
     * +
     */
    @Override
    protected void parsePlus() {
        var formula = (Formula) stack.pop();
        var plus = new Plus(formula);
        plus.setSheetName(sheetName);
        plus.setSheetIndex(sheetIndex);
        graph.addNode(plus);
        stack.push(plus);
    }

    /**
     * -
     */
    @Override
    protected void parseMinus() {
        var formula = (Formula) stack.pop();
        var minus = new Minus(formula);
        setOwnProperty(minus);
        graph.addNode(minus);
        stack.push(minus);
    }
//Unary END


//Union & Intersection BEGIN

    /**
     * F F
     */
    @Override
    protected void parseIntersection() {
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
    protected void parseUnion() {
        var rFormula = (Formula) stack.pop();
        var lFormula = (Formula) stack.pop();
        var union = new Union(lFormula, rFormula);
        setOwnProperty(union);
        graph.add(union);
        stack.push(union);
    }
//Union & Intersection END

//Constants BEGIN

    @Override
    protected void parseErrorLiteral(ERROR term) {
        setOwnProperty(term);
        err(term.toString(), rowFormula, colFormula);
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void parseBooleanLiteral(BOOL term) {
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void parseStringLiteral(TEXT term) {
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void parseIntLiteral(INT term) {
        graph.addNode(term);
        stack.push(term);
    }

    @Override
    protected void parseFloatLiteral(FLOAT term) {
        graph.addNode(term);
        stack.push(term);
    }

//Constants END
}
