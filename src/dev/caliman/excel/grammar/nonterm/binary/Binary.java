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

package dev.caliman.excel.grammar.nonterm.binary;

import dev.caliman.excel.grammar.annotations.NonTerminal;
import dev.caliman.excel.grammar.annotations.Production;
import dev.caliman.excel.grammar.lexicaltokens.CELL;
import dev.caliman.excel.grammar.nonterm.Formula;
import dev.caliman.excel.grammar.nonterm.FunctionCall;
import dev.caliman.excel.grammar.nonterm.ParenthesisFormula;
import dev.caliman.excel.grammar.nonterm.unary.Unary;
import org.jetbrains.annotations.NotNull;

/**
 * FunctionCall ::= Formula BinOp Formula
 * BinOp ::= + | - | * | / | ^ | < | > | = | <= | >= | <>
 *
 * Binary ::= Add | Sub | Mult | Div | Power | Lt | Gt | Eq | Leq | GtEq | Neq
 *
 * @author Massimo Caliman
 */
@NonTerminal
@Production(symbol = "Binary", expression = "Add")
@Production(symbol = "Binary", expression = "Sub")
@Production(symbol = "Binary", expression = "Mult")
@Production(symbol = "Binary", expression = "Div")
@Production(symbol = "Binary", expression = "Lt")
@Production(symbol = "Binary", expression = "Gt")
@Production(symbol = "Binary", expression = "Eq")
@Production(symbol = "Binary", expression = "Leq")
@Production(symbol = "Binary", expression = "GtEq")
@Production(symbol = "Binary", expression = "Neq")
public abstract class Binary extends FunctionCall {

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
        return operandToFormula(lFormula) + op + operandToFormula(rFormula);
    }

    private String operandToFormula(Formula operand) {
        if(operand instanceof CELL || operand instanceof Unary) return operand.id();
        else if(operand instanceof ParenthesisFormula)
            return operandToFormulaParenthesisFormula((ParenthesisFormula) operand);
        else return operand.toString();
    }

    private String operandToFormulaParenthesisFormula(ParenthesisFormula operand) {
        return operand.getFormula() instanceof Binary ?
                "(" + operand.getFormula().toString() + ")" :
                "(" + operand.getFormula().getAddress(false) + ")";
    }

    public Formula getlFormula() {
        return lFormula;
    }

    public Formula getrFormula() {
        return rFormula;
    }

}