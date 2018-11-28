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

import excel.grammar.formula.reference.CELL_REFERENCE;
import excel.grammar.formula.reference.RANGE;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.ptg.Area3DPxg;
import org.apache.poi.ss.formula.ptg.AreaPtg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.List;

public class RangeInternal {

    private final SpreadsheetVersion SPREADSHEET_VERSION = SpreadsheetVersion.EXCEL2007;
    private final Workbook workbook;
    private final Sheet sheet;

    private final int firstRow;
    private final int firstColumn;

    private final int lastRow;
    private final int lastColumn;
    RANGE tRANGE;
    private String sheetName;

    RangeInternal(Workbook workbook, Sheet sheet, AreaPtg t) {
        firstRow = t.getFirstRow();
        firstColumn = t.getFirstColumn();

        lastRow = t.getLastRow();
        lastColumn = t.getLastColumn();

        CELL_REFERENCE first = new CELL_REFERENCE(firstRow, firstColumn);
        CELL_REFERENCE last = new CELL_REFERENCE(lastRow, lastColumn);
        tRANGE = new RANGE(first, last);

        this.workbook = workbook;
        this.sheet = sheet;
        init();
    }

    RangeInternal(Workbook workbook, String sheetnamne, Area3DPxg t) {
        firstRow = t.getFirstRow();
        firstColumn = t.getFirstColumn();
        sheetName = sheetnamne;


        lastRow = t.getLastRow();
        lastColumn = t.getLastColumn();

        CELL_REFERENCE first = new CELL_REFERENCE(firstRow, firstColumn);
        CELL_REFERENCE last = new CELL_REFERENCE(lastRow, lastColumn);
        tRANGE = new RANGE(first, last);

        this.workbook = workbook;
        this.sheet = null;
        String refs = tRANGE.toString();

        AreaReference area = new AreaReference(sheetnamne + "!" + refs, SPREADSHEET_VERSION);
        List<Cell> cells = fromRange(area);

        for (Cell cell : cells)
            if (cell != null) {
                CellInternal excelType = new CellInternal(cell);
                tRANGE.add(excelType.valueOf());
            }
    }

    private void init() {
        String refs = tRANGE.toString();
        List<Cell> cells = range(refs);
        for (Cell cell : cells)
            if (cell != null) {
                CellInternal excelType = new CellInternal(cell);
                tRANGE.add(excelType.valueOf());
            }
    }

    private List<Cell> range(String refs) {
        AreaReference area = new AreaReference(sheet.getSheetName() + "!" + refs, SPREADSHEET_VERSION);
        return fromRange(area);
    }

    private List<Cell> fromRange(AreaReference area) {
        List<Cell> cells = new ArrayList<>();
        org.apache.poi.ss.util.CellReference[] cels = area.getAllReferencedCells();
        for (org.apache.poi.ss.util.CellReference cel : cels) {
            XSSFSheet ss = (XSSFSheet) workbook.getSheet(cel.getSheetName());
            Row r = ss.getRow(cel.getRow());
            if (r == null) continue;
            Cell c = r.getCell(cel.getCol());
            cells.add(c);
        }
        return cells;
    }

    private int getFirstRow() {
        return firstRow;
    }

    private int getFirstColumn() {
        return firstColumn;
    }

    private int getLastRow() {
        return lastRow;
    }

    private int getLastColumn() {
        return lastColumn;
    }

    String getSheetName() {
        return sheetName;
    }

    RANGE getRANGE() {
        return tRANGE;
    }

}
