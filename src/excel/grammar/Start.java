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

import excel.parser.AbstractParser;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author Massimo Caliman
 */
public abstract class Start {

    protected String sheetName;
    private int row;
    private int column;
    private String comment;
    private int sheetIndex;

    protected static String format(String str) {
        if (str == null) return "";
        return "\"" + str + "\"";
    }

    protected static String format(Boolean bool) {
        if (bool == null) return "";
        return bool ? "TRUE" : "FALSE";
    }

    protected static String format(Integer integer) {
        if (integer == null) return "";
        return integer.toString();
    }

    protected static String format(Double doub) {
        if (doub == null) return "";
        return doub.toString();
    }

    protected static String format(final Date date) {
        SimpleDateFormat format = new SimpleDateFormat("dd/MM/yyyy");
        return format.format(date);
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

    private String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String getAddress() {
        return AbstractParser.HelperInternal.cellAddress(getRow(), getColumn(), sheetName);
    }

    public String getAddress(boolean sheet) {
        return sheet ? AbstractParser.HelperInternal.cellAddress(getRow(), getColumn(), sheetName) : AbstractParser.HelperInternal.cellAddress(getRow(), getColumn());
    }

    @Override
    public int hashCode() {
        int hash = 5;
        hash = 53 * hash + row;
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

    private boolean sameAddr(Object obj) {
        final Start that = (Start) obj;
        return this.column == that.column && this.row == that.row && this.sheetIndex == that.sheetIndex;
    }

    @Override
    public boolean equals(final Object obj) {
        if (!(obj instanceof Start)) return false;
        final Start that = (Start) obj;
        if (this.row == -1 || that.row == -1)
            return (this.column == that.column && this.row == that.row && this.sheetIndex == that.sheetIndex);
        else return this.getAddress().equalsIgnoreCase(that.getAddress());
    }

    @Override
    public String toString() {
        return toString(true);
    }

    public abstract String toString(boolean address);

    public boolean test(String text) {
        return this.toString(true).equals(text);
    }

}
