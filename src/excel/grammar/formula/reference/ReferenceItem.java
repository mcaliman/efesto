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

/**
 * @author Massimo Caliman
 */
public class ReferenceItem extends Reference {

    public String value;

    int firstRow;
    int firstColumn;
    int lastRow;
    int lastColumn;


    public ReferenceItem() {
    }

    public ReferenceItem(String value) {
        this.value = value;
    }


    @Override
    public String toString() {
        return value;
    }

    public String toString(boolean address) {
        return format(this, address);
    }

    private String format(ReferenceItem referenceItem, boolean address) {
        StringBuilder buff = new StringBuilder();
        if (address) {
            if (this.sheetName != null && this.sheetName.length() > 0) buff.append(sheetName + "!");
            buff.append(referenceItem.getReferenceItemValue());
            buff.append(" = ");
            buff.append(referenceItem.values());
        } else {
            if (this.sheetName != null && this.sheetName.length() > 0) buff.append(sheetName + "!");
            buff.append(referenceItem.getReferenceItemValue());
        }
        return buff.toString();
    }


    public String getReferenceItemValue() {
        return value;
    }


    boolean horizzontal_range() {
        return firstRow == lastRow && firstColumn != lastColumn;
    }

    boolean vertical_range() {
        return firstColumn == lastColumn && firstRow != lastRow;
    }

    public String values() {
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
            //Cicla per riga
            StringBuilder buff = new StringBuilder();
            buff.append("[");
            int index = 0;
            for (int row = firstRow; row <= lastRow; row++) {
                buff.append("[");
                for (int col = firstColumn; col <= lastColumn; col++) {
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
    }


    public void setFirstRow(int firstRow) {
        this.firstRow = firstRow;
    }

    public void setFirstColumn(int firstColumn) {
        this.firstColumn = firstColumn;
    }

    public void setLastRow(int lastRow) {
        this.lastRow = lastRow;
    }

    public void setLastColumn(int lastColumn) {
        this.lastColumn = lastColumn;
    }
}
