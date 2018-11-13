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

import excel.grammar.Formula;
import excel.grammar.Start;

/**
 * @author Massimo Caliman
 */
public class ConstantArray extends Formula {
    private Object[][] array;

    public ConstantArray(Object[][] array) {
        this.array = array;
    }

    private Object[][] getArray() {
        return array;
    }

    @Override
    public String toString() {
        return format(this);
    }

    @Override
    public String toString(boolean address) {
        return format(this);
    }

    private String format(ConstantArray e) {
        Object[][] array = e.getArray();
        StringBuilder str = new StringBuilder();
        str.append(varname(e)).append(" = ");
        str.append("{");
        for (int i = 0; i < array.length; i++) {
            Object[] internal = array[i];
            str.append(internal[0]).append(",");
        }
        if (str.charAt(str.length() - 1) == ',') str.deleteCharAt(str.length() - 1);
        str.append("}");
        return str.toString();
    }

    private String varname(Start start) {
        return start.getAddr();
    }
}
