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

package excel.grammar.formula.functioncall.unary;

import excel.grammar.Formula;
import excel.grammar.formula.FunctionCall;
import excel.grammar.formula.reference.CELL_REFERENCE;
import org.jetbrains.annotations.NotNull;

/**
 * @author Massimo Caliman
 */
public abstract class Unary extends FunctionCall {

    private final Formula formula;
    private final String unOpPrefix;

    Unary(String unOpPrefix, Formula formula) {
        this.unOpPrefix = unOpPrefix;
        this.formula = formula;
    }

    @NotNull
    @Override
    public String toString() {
        return unOpPrefix + formula.toString();
    }

    @NotNull
    @Override
    public String toString(boolean address) {
        if (formula instanceof CELL_REFERENCE) {
            return address ?
                    getAddress(true) + " = " + unOpPrefix + ((CELL_REFERENCE) formula).getValue() :
                    unOpPrefix + ((CELL_REFERENCE) formula).getValue();
        } else {
            return getAddress() + " = " + unOpPrefix + formula.toString();
        }
    }


    public String getUnOpPrefix() {
        return unOpPrefix;
    }

}
