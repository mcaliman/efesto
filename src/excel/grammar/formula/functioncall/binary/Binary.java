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

import excel.ToFormula;
import excel.grammar.Formula;
import excel.grammar.formula.FunctionCall;
import excel.grammar.formula.ParenthesisFormula;
import excel.grammar.formula.functioncall.unary.Unary;
import excel.grammar.formula.reference.CELL_REFERENCE;
import org.jetbrains.annotations.NotNull;

import static excel.grammar.Grammar.closeparen;
import static excel.grammar.Grammar.openparen;

/**
 * @author Massimo Caliman
 */
public abstract class Binary extends FunctionCall implements ToFormula {

    private final String op;
    private final Formula lFormula;
    private final Formula rFormula;

    Binary(Formula lFormula, String op, Formula rFormula) {
        this.lFormula = lFormula;
        this.op = op;
        this.rFormula = rFormula;
    }

    @NotNull
    @Override
    public String toString() {
        return getAddress(true) + " = " + operandTo(lFormula) + op + operandTo(rFormula);
    }

    @NotNull
    public String toString(boolean address) {
        return address ?
                getAddress(true) + " = " + operandTo(lFormula) + op + operandTo(rFormula) :
                operandTo(lFormula) + op + operandTo(rFormula);
    }

    @Override
    public String toFormula() {
        return operandToFuctional(lFormula) + op + operandToFuctional(rFormula);
    }

    private String operandToFuctional(Formula operand) {
        if (operand instanceof CELL_REFERENCE || operand instanceof Unary) {
            return operand.id();
        } else if (operand instanceof ParenthesisFormula) {
            return operandToFunctional((ParenthesisFormula) operand);
        } else {
            return operand.toFormula();
        }
    }

    private String operandToFunctional(ParenthesisFormula operand) {
        return operand.getFormula() instanceof Binary ?
                "" + openparen + operand.getFormula().toFormula() + closeparen :
                "" + openparen + operand.getFormula().getAddress(false) + closeparen;
    }


    public Formula getlFormula() {
        return lFormula;
    }

    public Formula getrFormula() {
        return rFormula;
    }

    private String operandTo(Formula operand) {
        if (operand instanceof CELL_REFERENCE) {
            return operand.getAddress();
        } else if (operand instanceof ParenthesisFormula) {
            return operandTo((ParenthesisFormula) operand);
        } else if (operand instanceof Unary) {
            return ((Unary) operand).getUnOpPrefix() + operand.getAddress();
        } else {
            return operand.toString();
        }
    }

    private String operandTo(ParenthesisFormula operand) {
        return operand.getFormula() instanceof Binary ?
                openparen + operand.getFormula().toString(false) + closeparen :
                openparen + operand.getFormula().getAddress(true) + closeparen;
    }

}
