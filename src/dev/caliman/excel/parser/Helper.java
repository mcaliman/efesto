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

package dev.caliman.excel.parser;


import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ss.usermodel.Cell.*;

class Helper {


    private final Workbook workbook;

    public Helper(Workbook workbook) {
        this.workbook = workbook;

    }

    @Nullable
    public static Object valueOf(@Nullable Cell cell) {
        if(cell == null) return null;
        if(Helper.isDataType(cell))
            return cell.getDateCellValue();
        switch(cell.getCellType()) {
            case CELL_TYPE_STRING:
            case CELL_TYPE_BLANK:
                return cell.getStringCellValue();
            case CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case CELL_TYPE_FORMULA:
                if(cell.toString() != null && cell.toString().equalsIgnoreCase("TRUE")) {
                    return true;
                }
                if(cell.toString() != null && cell.toString().equalsIgnoreCase("FALSE")) {
                    return false;
                }
                return cell.toString();
            default:
                return null;
        }
    }

    private static boolean isDataType(Cell cell) {
        return cell.getCellType() == CELL_TYPE_NUMERIC && HSSFDateUtil.isCellDateFormatted(cell);
    }


    @NotNull
    public List<Cell> fromRange(@NotNull AreaReference area) {
        List<Cell> cells = new ArrayList<>();
        org.apache.poi.ss.util.CellReference[] cels = area.getAllReferencedCells();
        for(org.apache.poi.ss.util.CellReference cel : cels) {
            XSSFSheet ss = (XSSFSheet) workbook.getSheet(cel.getSheetName());
            Row r = ss.getRow(cel.getRow());
            if(r == null) continue;
            Cell c = r.getCell(cel.getCol());
            cells.add(c);
        }
        return cells;
    }


}
