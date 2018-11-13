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

package excel.grammar.formula.functioncall.binary;

import excel.grammar.Formula;
import excel.grammar.Start;
import excel.grammar.formula.FunctionCall;
import excel.grammar.formula.ParenthesisFormula;
import excel.grammar.formula.functioncall.unary.Unary;
import excel.grammar.formula.reference.CELL;

/**
 * @author Massimo Caliman
 */
public abstract class Binary extends FunctionCall {

    private final String op;
    private Formula lFormula;
    private Formula rFormula;

    Binary(Formula lFormula, String op, Formula rFormula) {
        this.lFormula = lFormula;
        this.op = op;
        this.rFormula = rFormula;
    }

    @Override
    public String toString() {
        return toString(true);
    }

    public String toString(boolean address) {
        if (address) return this.getAddr(true) + " = " + operandTo(lFormula) + op + operandTo(rFormula);
        else return operandTo(lFormula) + op + operandTo(rFormula);

    }

    public String getValue() {
        return operandTo(lFormula) + op + operandTo(rFormula);
    }

    public String getOp() {
        return op;
    }


    public Start getlFormula() {
        return lFormula;
    }


    public Formula getrFormula() {
        return rFormula;
    }


    private String operandTo(Start operand) {
        if (operand instanceof CELL) return operand.getAddr();
        else if (operand instanceof ParenthesisFormula) return format((ParenthesisFormula) operand, false);
        else if (operand instanceof Unary) return ((Unary) operand).getUnOpPrefix() + operand.getAddr();
        else return operand.toString();
    }

    private String format(ParenthesisFormula start, boolean address) {
        if (start.getFormula() instanceof Binary) return "(" + start.getFormula().toString(false) + ")";
        else return "(" + start.getFormula().getAddr(true) + ")";
    }
}
