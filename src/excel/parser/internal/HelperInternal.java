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

package excel.parser.internal;

public class HelperInternal {


    public static String reference(final int firstRow, final int firstCol, boolean isFirstRowRel, boolean isFirstColRel,
                                   int lastRow, int lastCol, boolean isLastRowRel, boolean isLastColRel
    ) {
        return cellAddress(firstRow, firstCol, isFirstRowRel, isFirstColRel) + ":" + HelperInternal.cellAddress(lastRow, lastCol, isLastRowRel, isLastColRel);
    }

    public static String columnAsLetter(final int column) {
        return org.apache.poi.ss.util.CellReference.convertNumToColString(column);
    }

    public static String cellAddress(final int row, final int column) {
        String letter = columnAsLetter(column);
        return (letter + (row + 1));
    }

    private static String cellAddress(final int row, final int column, boolean rowrelative, boolean columnrelative) {
        String letter = columnAsLetter(column);
        return (letter + (row + 1));
    }

    public static String cellAddress(final int row, final int column, final String sheetName) {
        StringBuilder buffer = new StringBuilder();
        if (sheetName != null)
            buffer.append(sheetName).append("!");
        buffer.append(cellAddress(row, column));
        return buffer.toString();
    }

}
