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

package excel.grammar;

import excel.ToFormula;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.text.SimpleDateFormat;
import java.util.Date;

import static excel.grammar.Grammar.epsilon;
import static excel.grammar.Grammar.exclamationmark;

/**
 * @author Massimo Caliman
 */
public abstract class Start implements ToFormula {

    private final static SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("dd/MM/yyyy");

    protected String sheetName;
    protected boolean singleSheet;
    private int row;
    private int column;
    private String comment;
    private int sheetIndex;

    @NotNull
    protected static String format(@Nullable final Date date) {
        return date == null ? epsilon : DATE_FORMAT.format(date);
    }

    public static String cellAddress(final int row, final int column, @Nullable final String sheetName) {
        StringBuilder buffer = new StringBuilder();
        if (sheetName != null)
            buffer.append(sheetName).append(exclamationmark);
        buffer.append(cellAddress(row, column));
        return buffer.toString();
    }

    public static String cellAddress(final int row, final int column) {
        String letter = columnAsLetter(column);
        return (letter + (row + 1));
    }

    public static String columnAsLetter(int col) {
        int excelColNum = col + 1;
        StringBuilder colRef = new StringBuilder(2);
        int colRemain = excelColNum;
        while (colRemain > 0) {
            int thisPart = colRemain % 26;
            if (thisPart == 0) {
                thisPart = 26;
            }
            colRemain = (colRemain - thisPart) / 26;
            char colChar = (char) (thisPart + 64);
            colRef.insert(0, colChar);
        }
        return colRef.toString();
    }

    public void setSingleSheet(boolean singleSheet) {
        this.singleSheet = singleSheet;
    }

    public boolean isTerminal() {
        return false;
    }

    public String getComment() {
        return comment;
    }

    public void setComment(String comment) {
        this.comment = comment;
    }

    protected int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    protected int getColumn() {
        return column;
    }

    public void setColumn(int column) {
        this.column = column;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    @NotNull
    public String getAddress() {
        return Start.cellAddress(getRow(), getColumn(), sheetName);
    }

    @NotNull
    public String getAddress(boolean sheet) {
        return sheet ? Start.cellAddress(getRow(), getColumn(), sheetName) : cellAddress(getRow(), getColumn());
    }

    public String id() {
        return this.singleSheet ? cellAddress(getRow(), getColumn()) : sheetName + "!" + cellAddress(getRow(), getColumn());
    }

    @Override
    public int hashCode() {
        int hash = 5;
        hash = 53 * hash + this.row;
        hash = 53 * hash + this.column;
        hash = 53 * hash + this.sheetIndex;
        return hash;
    }

    protected boolean isArea() {
        return this.row == -1 && this.column == -1;
    }

    public void setAsArea() {
        this.column = -1;
        this.row = -1;
    }

    @Override
    public boolean equals(final Object obj) {
        if (!(obj instanceof Start)) return false;
        final Start that = (Start) obj;
        if (this.row == -1 || that.row == -1)
            return (this.column == that.column && this.row == that.row && this.sheetIndex == that.sheetIndex);
        else return this.getAddress().equalsIgnoreCase(that.getAddress());
    }

    @Nullable
    @Override
    public String toString() {
        return toString(true);
    }

    @Nullable
    public abstract String toString(boolean address);

    /*public boolean test(String text) {
        return this.toString(true).equals(text);
    }*/

    public boolean testToFunctional(String text) {
        return (this.id() + " = " + this.toFormula()).equals(text);
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }


}
