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

package excel.grammar.formula.functioncall;

import excel.ToFunctional;
import excel.grammar.Formula;
import excel.grammar.formula.FunctionCall;
import excel.grammar.formula.reference.CELL_REFERENCE;
import excel.grammar.formula.reference.NamedRange;
import excel.grammar.formula.reference.RangeReference;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import static excel.grammar.Grammar.closeparen;
import static excel.grammar.Grammar.openparen;

/**
 * @author Massimo Caliman
 */
public abstract class EXCEL_FUNCTION extends FunctionCall implements ToFunctional {

    protected Formula[] args;

    protected EXCEL_FUNCTION(Formula... args) {
        this.args = args;
    }

    public Formula[] getArgs() {
        return args;
    }

    @NotNull
    @Override
    public String toString() {
        return getAddress() + " = " + getName() + openparen + argumentsToString() + closeparen;
    }

    @NotNull
    public String toString(boolean address) {
        return address ?
                getAddress() + " = " + getName() + openparen + argumentsToString() + closeparen :
                getName() + openparen + argumentsToString() + closeparen;
    }

    @Override
    public String toFunctional() {
        return getName() + openparen + argumentsToFunctional() + closeparen;
    }

    private String getName() {
        return getClass().getSimpleName();
    }

    private String argumentsToString() {
        var buff = new StringBuilder();
        Formula[] args = getArgs();
        if (args == null || args.length == 0) return "Missing";
        for (Formula arg : args) buff.append(argumentToString(arg)).append(",");
        if (buff.charAt(buff.length() - 1) == ',') buff.deleteCharAt(buff.length() - 1);
        return buff.toString();
    }

    private String argumentToString(@Nullable Formula operand) {
        if (operand == null) return "Missing";
        return operand instanceof CELL_REFERENCE ? operand.getAddress() : operand.toString(false);
    }

    private String argumentsToFunctional() {
        var buff = new StringBuilder();
        Formula[] args = getArgs();
        if (args == null || args.length == 0) return "Missing";
        for (Formula arg : args) buff.append(argumentToFunctional(arg)).append(",");
        if (buff.charAt(buff.length() - 1) == ',') buff.deleteCharAt(buff.length() - 1);
        return buff.toString();
    }

    private String argumentToFunctional(@Nullable Formula operand) {
        if (operand == null) return "Missing";
        //System.err.println(operand.getClass().getSimpleName());
        if (operand instanceof RangeReference || operand instanceof NamedRange) {
            return operand.id();
        } else if (operand instanceof CELL_REFERENCE) {
            return operand.id();
        } else {
            return operand.toFunctional();
        }
    }

}
