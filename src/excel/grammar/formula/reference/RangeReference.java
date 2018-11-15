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
import excel.grammar.formula.Reference;

/**
 * @author Massimo Caliman
 */
public final class RangeReference extends Reference {

    private final CELL reference1;
    private final CELL reference2;

    public RangeReference(CELL reference1, CELL reference2) {
        this.reference1 = reference1;
        this.reference2 = reference2;
    }

    private boolean horizzontal_range() {
        return reference1.getRow() == reference2.getRow() && reference1.getColumn() != reference2.getColumn();
    }

    private boolean vertical_range() {
        return reference1.getColumn() == reference2.getColumn() && reference1.getRow() != reference2.getRow();
    }

    private String values() {
        try {
            if (horizzontal_range() || vertical_range()) {
                StringBuilder buff = new StringBuilder();
                buff.append("[ ");
                for (Object val : vals)
                    if (val instanceof String)
                        buff.append("\"").append(val).append("\" ");
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
                            buff.append("\"").append(vals.get(index)).append("\" ");
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

    @Override
    public String toString() {
        return "" + this.reference1.getAddress() + ":" + this.reference2.getAddress();
    }

    @Override
    public String toString(boolean address) {
        return format(this, address);
    }


    private String format(RangeReference e, boolean address) {
        CELL ref1 = e.getReference1();
        CELL ref2 = e.getReference2();
        String addr1 = varname(ref1);
        String addr2 = varname(ref2);
        StringBuilder buff = new StringBuilder();
        if (address && !e.sameAddr(ref1) && e.getColumn() != -1 && e.getRow() != -1)
            buff.append(varname(e)).append(" = ");

        String sheetName = e.getSheetName();
        if (sheetName != null && sheetName.trim().length() > 0) buff.append(sheetName).append("!");

        buff.append(addr1).append(":").append(addr2);
        if (address) {
            String values = e.values();
            buff.append(" = ").append(values);
        }
        return buff.toString();
    }

    private String varname(Start start) {
        return start.getAddress();
    }

    private CELL getReference1() {
        return reference1;
    }

    private CELL getReference2() {
        return reference2;
    }

}
