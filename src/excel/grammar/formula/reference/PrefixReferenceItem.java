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
 * PrefixReferenceItem ::= ⟨Prefix⟩ ⟨ReferenceItem⟩
 * <p>
 * ⟨Prefix⟩ ::= SHEET
 * | ‘’’ SHEET-QUOTED
 * | ⟨File⟩ SHEET
 * | ‘’’ ⟨File⟩ SHEET-QUOTED
 * | FILE ‘!’
 * | MULTIPLE-SHEETS
 * | ⟨File⟩ MULTIPLE-SHEETS
 *
 * @author Massimo Caliman
 */
public final class PrefixReferenceItem extends Reference {

    private final Prefix prefix;

    private final String reference;

    private RANGE tRANGE;

    private int firstRow;
    private int firstColumn;

    private int lastRow;
    private int lastColumn;

    public PrefixReferenceItem(Prefix prefix, String reference, RANGE tRANGE) {
        this.prefix = prefix;
        this.reference = reference;
        this.tRANGE = tRANGE;
        if (this.tRANGE != null) {
            setAsArea();
            add(tRANGE.values());
            setFirstRow(tRANGE.getFirst().getRow());
            setFirstColumn(tRANGE.getFirst().getColumn());
            setLastRow(tRANGE.getLast().getRow());
            setLastColumn(tRANGE.getLast().getColumn());
        }
    }

    /**
     * If not address required
     *
     * @return
     */
    @Override
    public String toString() {
        return prefix + reference;
    }

    public String toString(boolean address) {
        return address ? ifIsNotArea() + prefix + reference + " = " + values() : toString();

    }

    private String ifIsNotArea() {
        return !isArea() ? getAddress() + " = " : "";
    }

// --Commented out by Inspection START (23/11/2018 08:30):
//    public String getSheetName() {
//        return sheetName;
//    }
// --Commented out by Inspection STOP (23/11/2018 08:30)

    private boolean is_HORIZONTAL_RANGE() {
        return firstRow == lastRow && firstColumn != lastColumn;
    }

    private boolean is_VERTICAL_RANGE() {
        return firstColumn == lastColumn && firstRow != lastRow;
    }

    //@TODO simplify
    private String values() {
        if (vals.isEmpty()) return "[]";
        if (is_HORIZONTAL_RANGE() || is_VERTICAL_RANGE()) {
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

    private void setFirstRow(int firstRow) {
        this.firstRow = firstRow;
    }

    private void setFirstColumn(int firstColumn) {
        this.firstColumn = firstColumn;
    }

    private void setLastRow(int lastRow) {
        this.lastRow = lastRow;
    }

    private void setLastColumn(int lastColumn) {
        this.lastColumn = lastColumn;
    }
}