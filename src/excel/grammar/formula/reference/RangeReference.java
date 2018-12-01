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

import excel.grammar.formula.Reference;

import static excel.grammar.Grammar.colon;
import static excel.grammar.Grammar.doublequote;
import static excel.grammar.Grammar.exclamationmark;

/**
 * @author Massimo Caliman
 */
public final class RangeReference extends Reference {

    private final CELL_REFERENCE reference1;
    private final CELL_REFERENCE reference2;

    public RangeReference(CELL_REFERENCE reference1, CELL_REFERENCE reference2) {
        this.reference1 = reference1;
        this.reference2 = reference2;
    }

    private boolean horizzontal_range() {
        return reference1.getRow() == reference2.getRow() && reference1.getColumn() != reference2.getColumn();
    }

    private boolean vertical_range() {
        return reference1.getColumn() == reference2.getColumn() && reference1.getRow() != reference2.getRow();
    }

    @Override
    public String toString() {
        return reference1.getAddress() + colon + reference2.getAddress();
    }

    @Override
    public String toString(boolean address) {
        return address ?
                sheetName + exclamationmark + reference1.getAddress() + colon + reference2.getAddress() + " = " + values() :
                sheetName + exclamationmark + reference1.getAddress() + colon + reference2.getAddress();
    }

    private String values() {
        try {
            if (horizzontal_range() || vertical_range()) {
                StringBuilder buff = new StringBuilder();
                buff.append("[ ");
                for (Object val : vals)
                    if (val instanceof String)
                        buff.append(doublequote).append(val).append("\" ");
                    else
                        buff.append(val).append(" ");
                if (buff.length() > 1)
                    buff.deleteCharAt(buff.length() - 1);
                buff.append(" ]");
                return buff.toString();
            } else {
                StringBuilder buff = new StringBuilder();
                buff.append("[");
                int index = 0;
                for (int row = reference1.getRow(); row <= reference2.getRow(); row++) {
                    buff.append("[");
                    for (int col = reference1.getColumn(); col <= reference2.getColumn(); col++) {
                        if (vals.get(index) instanceof String)
                            buff.append(doublequote).append(vals.get(index)).append("\" ");
                        else
                            buff.append(vals.get(index)).append(" ");
                        index++;
                    }
                    buff.deleteCharAt(buff.length() - 1);
                    buff.append("]");
                }
                buff.append("]");
                return buff.toString();
            }
        } catch (IndexOutOfBoundsException e) {
            return "IndexOutOfBoundsException";
        }
    }

}