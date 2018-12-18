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

package excel.grammar.formula;

import excel.ToFunctional;
import excel.grammar.Formula;
import excel.grammar.formula.functioncall.binary.Binary;

import static excel.grammar.Grammar.closeparen;
import static excel.grammar.Grammar.openparen;

/**
 * @author Massimo Caliman
 */
public final class ParenthesisFormula extends Formula implements ToFunctional {

    private final Formula formula;

    public ParenthesisFormula(Formula formula) {
        this.formula = formula;
    }

    @Override
    public String toString() {
        return openparen + formula.toString() + closeparen;
    }

    @Override
    public String toFunctional() {
        return openparen + formula.toFunctional() + closeparen;
    }

    public String toString(boolean address) {
        return isBinary() ?
                formula.toString(false) :
                formula.toString();
    }

    private boolean isBinary() {
        return formula instanceof Binary;
    }

    public Formula getFormula() {
        return formula;
    }

}
