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

import dev.caliman.excel.ToFormula;
import dev.caliman.excel.grammar.formula.Reference;
import org.jetbrains.annotations.Nullable;

/**
 * @author Massimo Caliman
 */
public final class PrefixReferenceItem extends Reference implements ToFormula {

    private final Prefix prefix;

    private final String reference;

    private int firstRow;
    private int firstColumn;

    private int lastRow;
    private int lastColumn;

    public PrefixReferenceItem(Prefix prefix, String reference, @Nullable RANGE tRANGE) {
        this.prefix = prefix;
        this.reference = reference;
        if (tRANGE != null) {
            setAsArea();
            add(tRANGE.values());
            setFirstRow(tRANGE.getFirst().getRow());
            setFirstColumn(tRANGE.getFirst().getColumn());
            setLastRow(tRANGE.getLast().getRow());
            setLastColumn(tRANGE.getLast().getColumn());
        }
    }

    @Override
    public String toFormula() {
        return isArea() ? values() : prefix + reference;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        PrefixReferenceItem that = (PrefixReferenceItem) o;

        return this.prefix.equals(that.prefix) &&
                this.reference.equals(that.reference) &&
                this.sheetName.equals(that.sheetName);
        //return that.toString(true).equals(this.toString(true));
    }

    public String id() {
        return !isArea() ? getAddress(!this.singleSheet) : prefix + reference;
    }

    private boolean is_HORIZONTAL_RANGE() {
        return firstRow == lastRow && firstColumn != lastColumn;
    }

    private boolean is_VERTICAL_RANGE() {
        return firstColumn == lastColumn && firstRow != lastRow;
    }

    private String values() {
        return values(firstRow, firstColumn, lastRow, lastColumn, vals, (is_HORIZONTAL_RANGE() || is_VERTICAL_RANGE()));
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