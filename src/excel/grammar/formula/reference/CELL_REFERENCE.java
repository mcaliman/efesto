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

import java.util.Date;

/**
 * @author mcaliman
 */
public final class CELL_REFERENCE extends ReferenceItem {

    private final int row;
    private final int column;

    private Object value;

    public CELL_REFERENCE(int row, int column) {
        this.row = row;
        this.column = column;
    }

    @Override
    public int getRow() {
        return row;
    }

    @Override
    public int getColumn() {
        return column;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    @Override
    public String toString() {
        if (value instanceof String) return getAddress() + " = " + format((String) value);
        else if (value instanceof Boolean) return getAddress() + " = " + format((Boolean) value);
        else if (value instanceof Integer) return getAddress() + " = " + format((Integer) value);
        else if (value instanceof Double) return getAddress() + " = " + format((Double) value);
        else if (value instanceof Date) return getAddress() + " = " + format((Date) value);
        else return null;
    }

    @Override
    public String toString(boolean address) {
        return toString();
    }

}
