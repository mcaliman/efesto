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

import excel.ToFunctional;
import excel.grammar.formula.Constant;
import excel.grammar.formula.constant.*;
import org.jetbrains.annotations.Nullable;

import java.util.Date;

/**
 * @author mcaliman
 */
public final class CELL_REFERENCE extends ReferenceItem implements ToFunctional {

    private final int row;
    private final int column;

    private Constant constant;

    public CELL_REFERENCE(int row, int column) {
        this.row = row;
        this.column = column;
    }

    public CELL_REFERENCE(int row, int column, String comment) {
        this.row = row;
        this.column = column;
        this.setComment(comment);
    }

    @Override
    public int getRow() {
        return row;
    }

    @Override
    public int getColumn() {
        return column;
    }

    public Constant getValue() {
        return constant;
    }

    public void setValue(Object value) {
        if (value instanceof String) constant = new TEXT((String) value);
        else if (value instanceof Boolean) constant = new BOOL((Boolean) value);
        else if (value instanceof Integer) constant = new INT((Integer) value);
        else if (value instanceof Double) constant = new FLOAT((Double) value);
        else if (value instanceof Date) constant = new DATETIME((Date) value);
    }

    @Override
    public String toString() {
        return getAddress() + " = " + constant.toString();
    }

    @Nullable
    @Override
    public String toString(boolean address) {
        return toString();
    }

    @Override
    public String toFunctional() {
        return constant != null ? constant.toString() : "null";
    }

}
