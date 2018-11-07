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

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.RichTextString;

import java.util.Date;

import static org.apache.poi.ss.usermodel.Cell.*;

/**
 * @author mcaliman
 */
public class CellInternal {

    private final Cell cell;
    private String comment;

    public CellInternal(Cell cell) {
        this.cell = cell;
        Comment cellComment = this.cell.getCellComment();
        comment = comment(cellComment);
        CellStyle style = this.cell.getCellStyle();
        String format = style.getDataFormatString();
    }

    public String getComment() {
        return comment;
    }

    private String comment(Comment comment) {
        if (comment == null) return null;
        RichTextString text = comment.getString();
        if (text == null) return null;
        return text.getString();

    }

    public Class type() {
        if (this.cell == null) return null;
        else if (isDate()) return Date.class;
        else if (isNumeric()) return Double.class;
        else if (isBoolean()) return Boolean.class;
        else if (isString()) return String.class;
        else return Object.class;
    }

    public Object valueOf() {
        if (cell == null) return null;
        if (isDataType(cell))
            return cell.getDateCellValue();
        switch (cell.getCellType()) {
            case CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case CELL_TYPE_BLANK:
                return cell.getStringCellValue();
            case CELL_TYPE_FORMULA:
                if (cell.toString() != null && cell.toString().equalsIgnoreCase("true")) {
                    return true;
                }
                if (cell.toString() != null && cell.toString().equalsIgnoreCase("false")) {
                    return false;
                }

                return cell.toString();
            default:
                return null;
        }
    }

    private boolean isDate() {
        return this.cell.getCellType() == CELL_TYPE_NUMERIC && HSSFDateUtil.isCellDateFormatted(this.cell);
    }

    private boolean isString() {
        return this.cell.getCellType() == CELL_TYPE_STRING;
    }

    private boolean isNumeric() {
        return this.cell.getCellType() == CELL_TYPE_NUMERIC;
    }

    private boolean isBoolean() {
        return this.cell.getCellType() == CELL_TYPE_BOOLEAN;
    }

    private Class typeOf(Cell c) {
        if (c == null) return null;
        if (isDataType(c)) return Date.class;
        switch (c.getCellType()) {
            case CELL_TYPE_STRING:
                return String.class;
            case CELL_TYPE_NUMERIC:
                return Double.class;
            case CELL_TYPE_BOOLEAN:
                return Boolean.class;
            default:
                return Object.class;
        }
    }

    private boolean isDataType(Cell c) {
        return c.getCellType() == CELL_TYPE_NUMERIC && HSSFDateUtil.isCellDateFormatted(c);
    }

    public Class internalFormulaResultType() {
        int type = cell.getCachedFormulaResultType();
        if (isDataType(cell))
            return Date.class;
        return internalFormulaResultType(type);
    }

    private Class internalFormulaResultType(int type) {
        switch (type) {
            case CELL_TYPE_STRING:
                return String.class;
            case CELL_TYPE_NUMERIC:
                return Double.class;
            case CELL_TYPE_BOOLEAN:
                return Boolean.class;
            default:
                return Object.class;
        }
    }

}
