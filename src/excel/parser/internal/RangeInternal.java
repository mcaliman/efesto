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

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.ptg.Area3DPxg;
import org.apache.poi.ss.formula.ptg.AreaNPtg;
import org.apache.poi.ss.formula.ptg.AreaPtg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.List;

/**
 * @author mcaliman
 */
public class RangeInternal {

    private final static SpreadsheetVersion SPREADSHEET_VERSION = SpreadsheetVersion.EXCEL2007;
    private final Workbook workbook;
    private final Sheet sheet;

    private final int firstRow;
    private final int firstColumn;

    private final boolean firstRowRelative;
    private final boolean firstColumnRelative;

    private final int lastRow;
    private final int lastColumn;

    private final boolean lastRowRelative;
    private final boolean lastColumnRelative;

    private List<Object> values;

    private String sheetName;

    public RangeInternal(Workbook workbook, Sheet sheet, AreaNPtg t) {
        firstRow = t.getFirstRow();
        firstColumn = t.getFirstColumn();

        firstRowRelative = t.isFirstRowRelative();
        firstColumnRelative = t.isFirstColRelative();

        lastRow = t.getLastRow();
        lastColumn = t.getLastColumn();

        lastRowRelative = t.isLastRowRelative();
        lastColumnRelative = t.isLastColRelative();
        this.workbook = workbook;
        this.sheet = sheet;
        init();
    }

    public RangeInternal(Workbook workbook, Sheet sheet, AreaPtg t) {
        firstRow = t.getFirstRow();
        firstColumn = t.getFirstColumn();

        firstRowRelative = t.isFirstRowRelative();
        firstColumnRelative = t.isFirstColRelative();

        lastRow = t.getLastRow();
        lastColumn = t.getLastColumn();

        lastRowRelative = t.isLastRowRelative();
        lastColumnRelative = t.isLastColRelative();
        this.workbook = workbook;
        this.sheet = sheet;
        init();
    }

    public RangeInternal(Workbook workbook, String sheetnamne, Area3DPxg t) {
        firstRow = t.getFirstRow();
        firstColumn = t.getFirstColumn();
        sheetName = sheetnamne;
        firstRowRelative = t.isFirstRowRelative();
        firstColumnRelative = t.isFirstColRelative();

        lastRow = t.getLastRow();
        lastColumn = t.getLastColumn();

        lastRowRelative = t.isLastRowRelative();
        lastColumnRelative = t.isLastColRelative();
        this.workbook = workbook;
        this.sheet = null;
        String refs = HelperInternal.reference(firstRow, firstColumn, firstRowRelative, firstColumnRelative, lastRow, lastColumn, lastRowRelative, lastColumnRelative);

        //List<Cell> cells = range(refs);

        AreaReference area = new AreaReference(sheetnamne + "!" + refs, SPREADSHEET_VERSION);
        List<Cell> cells = fromRange(area);


        values = new ArrayList<>();
        for (Cell cell : cells)
            if (cell != null) {
                CellInternal excelType = new CellInternal(cell);
                values.add(excelType.valueOf());
            }
    }

    private void init() {

        String refs = HelperInternal.reference(firstRow, firstColumn, firstRowRelative, firstColumnRelative, lastRow, lastColumn, lastRowRelative, lastColumnRelative);

        List<Cell> cells = range(refs);
        values = new ArrayList<>();
        for (Cell cell : cells)
            if (cell != null) {
                CellInternal excelType = new CellInternal(cell);
                values.add(excelType.valueOf());
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

    public int getFirstRow() {
        return firstRow;
    }

    public int getFirstColumn() {
        return firstColumn;
    }

    public boolean isFirstRowRelative() {
        return firstRowRelative;
    }

    public boolean isFirstColumnRelative() {
        return firstColumnRelative;
    }

    public int getLastRow() {
        return lastRow;
    }

    public int getLastColumn() {
        return lastColumn;
    }

    public boolean isLastRowRelative() {
        return lastRowRelative;
    }

    public boolean isLastColumnRelative() {
        return lastColumnRelative;
    }

    public List<Object> getValues() {
        return values;
    }

    public String getSheetName() {
        return sheetName;
    }
}
