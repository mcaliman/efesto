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

package dev.caliman.excel.parser.internal;

import dev.caliman.excel.grammar.formula.reference.CELL_REFERENCE;
import dev.caliman.excel.grammar.formula.reference.RANGE;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.ptg.Area3DPxg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.jetbrains.annotations.NotNull;

import java.util.List;

class RangeInternal {

    @NotNull
    private final RANGE tRANGE;
    private final String sheetName;


    RangeInternal(Workbook workbook, String sheetnamne, Area3DPxg t) {
        Helper helper = new Helper(workbook);
        int firstRow = t.getFirstRow();
        int firstColumn = t.getFirstColumn();
        sheetName = sheetnamne;
        int lastRow = t.getLastRow();
        int lastColumn = t.getLastColumn();

        CELL_REFERENCE first = new CELL_REFERENCE(firstRow, firstColumn);
        CELL_REFERENCE last = new CELL_REFERENCE(lastRow, lastColumn);
        tRANGE = new RANGE(first, last);
        String refs = tRANGE.toString();
        SpreadsheetVersion SPREADSHEET_VERSION = SpreadsheetVersion.EXCEL2007;
        AreaReference area = new AreaReference(sheetnamne + "!" + refs, SPREADSHEET_VERSION);
        List<Cell> cells = helper.fromRange(area);

        for (Cell cell : cells)
            if (cell != null) {
                tRANGE.add(Helper.valueOf(cell));
            }
    }


    String getSheetName() {
        return sheetName;
    }

    @NotNull
    RANGE getRANGE() {
        return tRANGE;
    }

}
