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

import excel.grammar.formula.constant.*;

import java.util.Date;

/**
 * @author mcaliman
 */
public final class CELL extends ReferenceItem {

    private final int row;
    private final int column;

    private Object value;

    public CELL(int row, int column, boolean rowrelative, boolean columnrelative) {
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
        return format(this);
    }

    @Override
    public String toString(boolean address) {
        return toString();
    }

    private String format(CELL e) {
        Object value = e.getValue();
        if (value instanceof String) return e.getAddr() + " = " + TEXT.format((String) value);
        else if (value instanceof Boolean) return e.getAddr() + " = " + BOOL.format((Boolean) value);
        else if (value instanceof Integer) return e.getAddr() + " = " + INT.format((Integer) value);
        else if (value instanceof Double) return e.getAddr() + " = " + FLOAT.format((Double) value);
        else if (value instanceof Date) return e.getAddr() + " = " + DATETIME.format((Date) value);
        return e.toString();
    }

}
