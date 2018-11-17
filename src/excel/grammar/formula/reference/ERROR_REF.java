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

package excel.grammar.formula.reference;

import excel.grammar.Start;

/**
 * @author Massimo Caliman
 */
public final class ERROR_REF extends ReferenceItem {

    public ERROR_REF() {
        super("#REF");
        System.err.println("ERROR-REF Reference error literal #REF!");
    }

    public String toString(boolean addr) {
        return format(this);

    }

    private String format(ERROR_REF e) {
        String buff = varname(e) + " = " +
                e.toString();
        return buff;
    }

    private String varname(Start start) {
        return start.getAddress();
    }

}
