/*
 * Efesto - Excel Formula Extractor System and Topological Ordering algorithm.
 * Copyright (C) 2017 Massimo Caliman mcaliman@gmail.com
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
 * please direct inquiries about Efesto licensing to mcaliman@gmail.com
 */

package dev.caliman.excel.grammar.formula.reference;

import dev.caliman.excel.grammar.annotations.LexicalTokens;
import dev.caliman.excel.grammar.formula.Constant;
import dev.caliman.excel.grammar.formula.constant.*;

import java.util.Date;

/**
 * @author mcaliman
 */
@LexicalTokens(name = "CELL", description = "Cell reference", content = "$? [A-Z]+ $? [0-9]+", priority = 2)
public final class CELL extends ReferenceItem {

    private final int row;
    private final int column;

    private Constant constant;

    public CELL(int row, int column) {
        this.row = row;
        this.column = column;
    }

    public int getRow() {
        return row;
    }

    public int getColumn() {
        return column;
    }

    public Constant getValue() {
        return constant;
    }

    public void setValue(Object value) {
        if(value instanceof String) constant = new TEXT((String) value);
        else if(value instanceof Boolean) constant = new BOOL((Boolean) value);
        else if(value instanceof Integer) constant = new INT((Integer) value);
        else if(value instanceof Double) constant = new FLOAT((Double) value);
        else if(value instanceof Date) constant = new DATE((Date) value);
    }

    @Override
    public String toString() {
        return constant != null ? constant.toString() : "null";
    }

}
