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
import excel.grammar.formula.Reference;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import static excel.grammar.Grammar.epsilon;

/**
 * @author Massimo Caliman
 */
public final class PrefixReferenceItem extends Reference implements ToFunctional {

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


    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        PrefixReferenceItem that = (PrefixReferenceItem) o;
        return that.toString(true).equals(this.toString(true));
    }

    public String id() {
        return this.singleSheet ?
                ifIsNotArea(false) + prefix + reference :
                ifIsNotArea() + prefix + reference
                ;
    }


    @Override
    public String toFunctional() {
        return values();
    }

    @NotNull
    private String ifIsNotArea(boolean address) {
        return !isArea() ? getAddress(address) + " = " : epsilon;
    }


    @NotNull
    private String ifIsNotArea() {
        return !isArea() ? getAddress() + " = " : epsilon;
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