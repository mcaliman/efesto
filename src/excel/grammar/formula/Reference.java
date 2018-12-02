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

import static excel.grammar.Grammar.doublequote;

/**
 * @author Massimo Caliman
 */
public abstract class Reference extends Formula {

    protected final List<Object> vals = new ArrayList<>();

    public void add(List<Object> values) {
        vals.addAll(values);
    }

    protected String values(int fRow,int fCol,int lRow,int lCol, List<Object> list,boolean isHorizzontalOrVerticalRange) {
        if (isHorizzontalOrVerticalRange) {
            StringBuilder buff = new StringBuilder();
            buff.append("[ ");
            for (Object val : list)
                if (val instanceof String)
                    buff.append(doublequote).append(val).append("\" ");
                else
                    buff.append(val).append(" ");
            if (buff.length() > 1)
                buff.deleteCharAt(buff.length() - 1);
            buff.append(" ]");
            return buff.toString();
        } else {
            //Cicla per riga
            StringBuilder buff = new StringBuilder();
            buff.append("[");
            int index = 0;
            for (int row = fRow; row <= lRow; row++) {
                buff.append("[");
                for (int col = fCol; col <= lCol; col++) {
                    if (list.get(index) instanceof String)
                        buff.append(doublequote).append(list.get(index)).append("\" ");
                    else
                        buff.append(list.get(index)).append(" ");
                    index++;
                }
                buff.deleteCharAt(buff.length() - 1);
                buff.append("]");
            }
            buff.append("]");
            return buff.toString();
        }
    }


}
