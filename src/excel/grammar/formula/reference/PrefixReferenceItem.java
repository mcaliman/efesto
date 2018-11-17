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
public final class PrefixReferenceItem extends Reference {

    private final Prefix prefix;

    private final String reference;

    private int firstRow;
    private int firstColumn;
    private int lastRow;
    private int lastColumn;

    public PrefixReferenceItem(Prefix prefix, String reference) {
        this.prefix = prefix;
        this.reference = reference;
    }

    @Override
    public String toString() {
        return prefix.toString() + reference;
    }

    public String toString(boolean address) {
        StringBuilder buff = new StringBuilder();
        if (address && !this.isArea()) buff.append((this).getAddress()).append(" = ");
        buff.append(prefix.toString()).append(this.reference);
        if (address) buff.append(" = ").append(values());
        return buff.toString();
    }

    public String getSheetName() {
        return sheetName;
    }

    private String getReference() {
        return reference;
    }

    private boolean horizzontal_range() {
        return firstRow == lastRow && firstColumn != lastColumn;
    }

    private boolean vertical_range() {
        return firstColumn == lastColumn && firstRow != lastRow;
    }

    private String values() {
        if (vals.isEmpty()) return "[]";

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