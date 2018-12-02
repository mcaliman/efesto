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

import java.util.ArrayList;
import java.util.List;

import static excel.grammar.Grammar.*;

/**
 * @author Massimo Caliman
 */
public abstract class Reference extends Formula {

    protected final List<Object> vals = new ArrayList<>();

    public void add(List<Object> values) {
        vals.addAll(values);
    }

    protected String values(int fRow, int fCol, int lRow, int lCol, List<Object> list, boolean isHorizzontalOrVerticalRange) {
        if (list.isEmpty()) return emptylist;
        if (isHorizzontalOrVerticalRange) {
            StringBuilder buff = new StringBuilder();
            buff.append(opensquareparen).append(space);
            for (Object val : list) buff.append(toString(val)).append(space);
            if (buff.length() > 1) buff.deleteCharAt(buff.length() - 1);
            buff.append(space).append(closesquareparen);
            return buff.toString();
        } else {
            StringBuilder buff = new StringBuilder();
            buff.append(opensquareparen);
            int index = 0;
            for (int row = fRow; row <= lRow; row++) {
                buff.append(opensquareparen);
                for (int col = fCol; col <= lCol; col++) {
                    buff.append(toString(list.get(index))).append(space);
                    index++;
                }
                buff.deleteCharAt(buff.length() - 1);
                buff.append(closesquareparen);
            }
            buff.append(closesquareparen);
            return buff.toString();
        }
    }

    protected String toString(Object value) {
        if (value instanceof String)
            return quote(value.toString());
        else
            return value.toString();
    }

}
