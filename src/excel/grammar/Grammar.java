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

    public Grammar() {
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

    public CELL cell(int row, int column) {
        return new CELL(row, column);
    }

    public EXCEL_FUNCTION builtinFunction(String name) throws UnsupportedBuiltinException {
        BuiltinFactory factory = new BuiltinFactory();
        factory.create(0, name);
        return (EXCEL_FUNCTION) factory.getBuiltInFunction();
    }


    public Reference as_reference(Start args) {
        if (args instanceof RangeReference) return (RangeReference) args;
        else if (args instanceof ReferenceItem) return (ReferenceItem) args;
        else if (args instanceof PrefixReferenceItem) return (PrefixReferenceItem) args;
        else return null;
    }

}
