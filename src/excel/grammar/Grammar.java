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

package excel.grammar;

import excel.grammar.formula.Reference;
import excel.grammar.formula.constant.*;
import excel.grammar.formula.functioncall.EXCEL_FUNCTION;
import excel.grammar.formula.functioncall.PercentFormula;
import excel.grammar.formula.functioncall.binary.*;
import excel.grammar.formula.functioncall.unary.Minus;
import excel.grammar.formula.functioncall.unary.Plus;
import excel.grammar.formula.reference.CELL;
import excel.grammar.formula.reference.PrefixReferenceItem;
import excel.grammar.formula.reference.RangeReference;
import excel.grammar.formula.reference.ReferenceItem;
import excel.parser.BuiltinFactory;
import excel.parser.UnsupportedBuiltinException;

public final class Grammar {

    public final static String plus = "+";
    public final static String minus = "-";
    public final static String concat = "&";
    public final static String add = "+";
    public final static String sub = "-";
    public final static String multiply = "*";
    public final static String div = "/";
    public final static String power = "^";
    public final static String intersection = " ";
    public final static String union = ",";
    public final static String percent = "%";
    public final static String eq = "=";
    public final static String lt = "<";
    public final static String gt = ">";
    public final static String leq = "<=";
    public final static String gteq = ">=";
    public final static String neq = "<>";

    public Grammar() {
    }

    public Eq eq(Start lExpr, Start rExpr) {
        return new Eq((Formula) lExpr, (Formula) rExpr);
    }

    public Lt lt(Start lExpr, Start rExpr) {
        return new Lt((Formula) lExpr, (Formula) rExpr);
    }

    public Gt gt(Start lExpr, Start rExpr) {
        return new Gt((Formula) lExpr, (Formula) rExpr);
    }

    public Leq leq(Start lExpr, Start rExpr) {
        return new Leq((Formula) lExpr, (Formula) rExpr);
    }

    public GtEq gteq(Start lExpr, Start rExpr) {
        return new GtEq((Formula) lExpr, (Formula) rExpr);
    }

    public Neq neq(Start lExpr, Start rExpr) {
        return new Neq((Formula) lExpr, (Formula) rExpr);
    }

    public Concat concat(Start lExpr, Start rExpr) {
        return new Concat((Formula) lExpr, (Formula) rExpr);
    }

    public Add add(Start lExpr, Start rExpr) {
        return new Add((Formula) lExpr, (Formula) rExpr);
    }

    public Sub subtrac(Start lExpr, Start rExpr) {
        return new Sub((Formula) lExpr, (Formula) rExpr);
    }

    public Mult multiply(Start lExpr, Start rExpr) {
        return new Mult((Formula) lExpr, (Formula) rExpr);
    }

    public Divide divide(Start lExpr, Start rExpr) {
        return new Divide((Formula) lExpr, (Formula) rExpr);
    }

    public Power power(Start lExpr, Start rExpr) {
        return new Power((Formula) lExpr, (Formula) rExpr);
    }

    public Intersection intersection(Start lExpr, Start rExpr) {
        return new Intersection((Formula) lExpr, (Formula) rExpr);
    }

    public Union union(Start lExpr, Start rExpr) {
        return new Union((Formula) lExpr, (Formula) rExpr);
    }

    public PercentFormula percentFormula(Start formula) {
        return new PercentFormula((Formula) formula);
    }

    public TEXT text(String value) {
        return new TEXT(value);
    }

    public FLOAT number(Double value) {
        return new FLOAT(value);
    }

    public INT number(Integer value) {
        return new INT(value);
    }

    public BOOL bool(Boolean value) {
        return new BOOL(value);
    }

    public ERROR error(String value) {
        return new ERROR(value);
    }

    public Plus plus(Start expr) {
        return new Plus((Formula) expr);
    }

    public Minus minus(Start expr) {
        return new Minus((Formula) expr);
    }

    public CELL cell(int row, int column, boolean rowRelative, boolean columnRelative) {
        return new CELL(row, column, rowRelative, columnRelative);
    }

    public EXCEL_FUNCTION builtinFunction(String name) throws UnsupportedBuiltinException {
        BuiltinFactory factory = new BuiltinFactory();
        factory.create(0, name);
        return (EXCEL_FUNCTION) factory.getBuiltInFunction();
    }

    public RangeReference rangeReference(int firstRow, int firstColumn,
                                         boolean isFirstRowRelative,
                                         boolean isFirstColRelative,
                                         int lastRow, int lastColumn, boolean isLastRowRelative, boolean isLastColRelative) {
        CELL firstCell = new CELL(firstRow, firstColumn, isFirstRowRelative, isFirstColRelative);
        CELL lastCell = new CELL(lastRow, lastColumn, isLastRowRelative, isLastColRelative);
        return new RangeReference(firstCell, lastCell);
    }


    public Reference as_reference(Start args) {
        if (args instanceof RangeReference) return (RangeReference) args;
        else if (args instanceof ReferenceItem) return (ReferenceItem) args;
        else if (args instanceof PrefixReferenceItem) return (PrefixReferenceItem) args;
        else return null;
    }

}
