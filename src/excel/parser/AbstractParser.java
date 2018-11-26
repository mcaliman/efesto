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

package excel.parser;

import excel.grammar.Start;
import excel.grammar.formula.reference.CELL_REFERENCE;
import excel.grammar.formula.reference.FILE;
import excel.grammar.formula.reference.RANGE;
import excel.grammar.formula.reference.SHEET;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.EvaluationName;
import org.apache.poi.ss.formula.EvaluationSheet;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.ptg.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;
import java.util.function.Predicate;
import java.util.stream.Stream;

import static org.apache.poi.ss.formula.ptg.ErrPtg.*;
import static org.apache.poi.ss.usermodel.Cell.*;

public abstract class AbstractParser {

    final boolean errors = false;
    private final String creator;
    private final String description;
    private final String keywords;
    private final String title;
    private final String subject;
    private final String category;
    private final String company;
    private final String template;
    private final String manager;
    private final Predicate<Ptg> arrayPtg = (Ptg t) -> t instanceof ArrayPtg;
    private final Predicate<Ptg> addPtg = (Ptg t) -> t instanceof AddPtg;
    private final Predicate<Ptg> area3DPxg = (Ptg t) -> t instanceof Area3DPxg;
    private final Predicate<Ptg> areaErrPtg = (Ptg t) -> t instanceof AreaErrPtg;
    private final Predicate<Ptg> areaPtg = (Ptg t) -> t instanceof AreaPtg;
    private final Predicate<Ptg> attrPtg = (Ptg t) -> t instanceof AttrPtg;
    private final Predicate<Ptg> boolPtg = (Ptg t) -> t instanceof BoolPtg;
    private final Predicate<Ptg> concatPtg = (Ptg t) -> t instanceof ConcatPtg;
    private final Predicate<Ptg> deleted3DPxg = (Ptg t) -> t instanceof Deleted3DPxg;
    private final Predicate<Ptg> deletedArea3DPtg = (Ptg t) -> t instanceof DeletedArea3DPtg;
    private final Predicate<Ptg> deletedRef3DPtg = (Ptg t) -> t instanceof DeletedRef3DPtg;
    private final Predicate<Ptg> dividePtg = (Ptg t) -> t instanceof DividePtg;
    private final Predicate<Ptg> equalPtg = (Ptg t) -> t instanceof EqualPtg;
    private final Predicate<Ptg> errPtg = (Ptg t) -> t instanceof ErrPtg;
    private final Predicate<Ptg> funcPtg = (Ptg t) -> t instanceof FuncPtg;
    private final Predicate<Ptg> funcVarPtg = (Ptg t) -> t instanceof FuncVarPtg;
    private final Predicate<Ptg> greaterEqualPtg = (Ptg t) -> t instanceof GreaterEqualPtg;
    private final Predicate<Ptg> greaterThanPtg = (Ptg t) -> t instanceof GreaterThanPtg;
    private final Predicate<Ptg> intersectionPtg = (Ptg t) -> t instanceof IntersectionPtg;
    private final Predicate<Ptg> intPtg = (Ptg t) -> t instanceof IntPtg;
    private final Predicate<Ptg> lessEqualPtg = (Ptg t) -> t instanceof LessEqualPtg;
    private final Predicate<Ptg> lessThanPtg = (Ptg t) -> t instanceof LessThanPtg;
    private final Predicate<Ptg> memErrPtg = (Ptg t) -> t instanceof MemErrPtg;
    private final Predicate<Ptg> missingArgPtg = (Ptg t) -> t instanceof MissingArgPtg;
    private final Predicate<Ptg> multiplyPtg = (Ptg t) -> t instanceof MultiplyPtg;
    private final Predicate<Ptg> namePtg = (Ptg t) -> t instanceof NamePtg;
    private final Predicate<Ptg> notEqualPtg = (Ptg t) -> t instanceof NotEqualPtg;
    private final Predicate<Ptg> numberPtg = (Ptg t) -> t instanceof NumberPtg;
    private final Predicate<Ptg> parenthesisPtg = (Ptg t) -> t instanceof ParenthesisPtg;
    private final Predicate<Ptg> percentPtg = (Ptg t) -> t instanceof PercentPtg;
    private final Predicate<Ptg> powerPtg = (Ptg t) -> t instanceof PowerPtg;
    private final Predicate<Ptg> ref3DPxg = (Ptg t) -> t instanceof Ref3DPxg;
    private final Predicate<Ptg> refErrorPtg = (Ptg t) -> t instanceof RefErrorPtg;
    private final Predicate<Ptg> refPtg = (Ptg t) -> t instanceof RefPtg;
    private final Predicate<Ptg> stringPtg = (Ptg t) -> t instanceof StringPtg;
    private final Predicate<Ptg> subtractPtg = (Ptg t) -> t instanceof SubtractPtg;
    private final Predicate<Ptg> unaryMinusPtg = (Ptg t) -> t instanceof UnaryMinusPtg;
    private final Predicate<Ptg> unaryPlusPtg = (Ptg t) -> t instanceof UnaryPlusPtg;
    private final Predicate<Ptg> unionPtg = (Ptg t) -> t instanceof UnionPtg;
    private final Predicate<Ptg> unknownPtg = (Ptg t) -> t instanceof UnknownPtg;
    private final Workbook workbook;
    private final XSSFEvaluationWorkbook evaluationWorkbook;
    public boolean verbose = false;
    public boolean metadata = false;
    private boolean workbookProtectionPresent;
    private String fileName;
    int formulaColumn;
    int formulaRow;
    int currentSheetIndex;
    String currentSheetName;
    private Sheet sheet;
    private EvaluationSheet evaluationSheet;

    AbstractParser(File file) throws InvalidFormatException, IOException {
        this(WorkbookFactory.create(file));
        this.fileName = file.getName();
    }

    private AbstractParser(Workbook workbook) {
        this.workbook = workbook;
        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) this.workbook;
        workbookProtectionPresent = xssfWorkbook.validateWorkbookPassword("password");
        POIXMLProperties props = xssfWorkbook.getProperties();
        POIXMLProperties.CoreProperties coreProperties = props.getCoreProperties();
        this.creator = coreProperties.getCreator();
        this.description = coreProperties.getDescription();
        this.keywords = coreProperties.getKeywords();
        this.title = coreProperties.getTitle();
        this.subject = coreProperties.getSubject();
        this.category = coreProperties.getCategory();
        POIXMLProperties.CustomProperties customProperties = props.getCustomProperties();
        customProperties.getProperty("Author");
        //List<CTProperty> list = customProperties.getUnderlyingProperties().getPropertyList();
        POIXMLProperties.ExtendedProperties extendedProperties = props.getExtendedProperties();
        this.company = extendedProperties.getUnderlyingProperties().getCompany();
        this.template = extendedProperties.getUnderlyingProperties().getTemplate();
        this.manager = extendedProperties.getUnderlyingProperties().getManager();
        this.evaluationWorkbook = XSSFEvaluationWorkbook.create((XSSFWorkbook) workbook);
        //System.out.println("Parse...");
    }

    public String getCreator() {
        return creator;
    }

    public String getDescription() {
        return description;
    }

    public String getKeywords() {
        return keywords;
    }

    public String getTitle() {
        return title;
    }

    public String getSubject() {
        return subject;
    }

    public String getCategory() {
        return category;
    }

    public String getCompany() {
        return company;
    }

    public String getTemplate() {
        return template;
    }

    public String getManager() {
        return manager;
    }

    public String getFileName() {
        return fileName;
    }

    public boolean isWorkbookProtectionPresent() {
        return workbookProtectionPresent;
    }

    void verbose(String text) {
        if (this.verbose) System.out.println(text);
    }

    void err(String string, int row, int column) {

    }

    void parse() {
        for (Sheet sht : this.workbook) {
            this.sheet = sht;
            this.workbookProtectionPresent = workbookProtectionPresent || ((XSSFSheet) sheet).validateSheetPassword("password");
            this.currentSheetIndex = this.workbook.getSheetIndex(this.sheet);
            this.currentSheetName = this.sheet.getSheetName();
            this.evaluationSheet = this.evaluationWorkbook.getSheet(this.currentSheetIndex);
            for (Row row : this.sheet)
                for (Cell cell : row)
                    if (cell != null) parse(cell);
                    else err("Cell is null.", formulaRow, formulaColumn);
        }
    }

    private void parse(Cell cell) {
        if (cell.getCellType() == CELL_TYPE_FORMULA)
            parseFormula(cell);
    }

    private void parseFormula(Cell cell) {
        verbose("Cell:" + cell.getClass().getSimpleName() + " " + cell.toString() + " " + cell.getCellType());
        CellInternal excelCell = new CellInternal(cell);
        String comment = excelCell.getComment();
        formulaColumn = cell.getColumnIndex();
        formulaRow = cell.getRowIndex();
        //noinspection unused
        Class internalFormulaResultTypeClass = excelCell.internalFormulaResultType();
        String formulaAddress = Start.cellAddress(formulaRow, formulaColumn);
        String formulaText = cell.getCellFormula();
        verbose(formulaAddress + " = " + formulaText);
        FormulaTokensInternal tokens = new FormulaTokensInternal(this.evaluationWorkbook, this.evaluationSheet);
        Ptg[] formulaPtgs = tokens.getFormulaTokens(formulaRow, formulaColumn);
        if (formulaPtgs == null) {
            System.err.println("ptgs empty or null for address " + formulaAddress);
            err("ptgs empty or null for address " + formulaAddress, formulaRow, formulaColumn);
            UDF(formulaText);
            return;
        }
        Start start = parse(formulaPtgs, formulaRow, formulaColumn);
        if (start != null) {
            start.setComment(comment);
            parseFormula(start);
        }
    }

    protected abstract void UDF(String arguments);

    private Start parse(Ptg[] ptgs, int row, int column) {
        parseFormulaInit();
        if (Ptg.doesFormulaReferToDeletedCell(ptgs)) doesFormulaReferToDeletedCell(row, column);
        for (Ptg ptg : ptgs) parse(ptg, row, column);
        return parseFormulaPost();
    }

    private void parse(Ptg p, int row, int column) {
        verbose("parse: " + p.getClass().getSimpleName());
        try (Stream<WhatIf> stream = Stream.of(
                new WhatIf(p, arrayPtg, (Ptg t) -> parseArrayPtg((ArrayPtg) t)),
                new WhatIf(p, addPtg, (Ptg t) -> add()),
                new WhatIf(p, area3DPxg, (Ptg t) -> parseArea3DPxg((Area3DPxg) t)),
                new WhatIf(p, areaErrPtg, (Ptg t) -> parseAreaErrPtg((AreaErrPtg) t)),
                new WhatIf(p, areaPtg, (Ptg t) -> parseAreaPtg((AreaPtg) t)),
                new WhatIf(p, attrPtg, (Ptg t) -> parseAttrPtg((AttrPtg) t)),
                new WhatIf(p, boolPtg, t -> BOOL(((BoolPtg) t).getValue())),
                new WhatIf(p, concatPtg, t -> concat()),
                new WhatIf(p, deleted3DPxg, (Ptg t) -> parseDeleted3DPxg((Deleted3DPxg) t)),
                new WhatIf(p, deletedArea3DPtg, (Ptg t) -> parseDeletedArea3DPtg((DeletedArea3DPtg) t)),
                new WhatIf(p, deletedRef3DPtg, (Ptg t) -> parseDeletedRef3DPtg((DeletedRef3DPtg) t)),
                new WhatIf(p, dividePtg, t -> div()),
                new WhatIf(p, equalPtg, t -> eq()),
                new WhatIf(p, errPtg, (Ptg t) -> parseErrPtg((ErrPtg) t)),
                new WhatIf(p, funcPtg, (Ptg t) -> parseFuncPtg((FuncPtg) t)),
                new WhatIf(p, funcVarPtg, (Ptg t) -> parseFuncVarPtg((FuncVarPtg) t)),
                new WhatIf(p, greaterEqualPtg, t -> gteq()),
                new WhatIf(p, greaterThanPtg, t -> gt()),
                new WhatIf(p, intersectionPtg, t -> intersection()),
                new WhatIf(p, intPtg, t -> INT(((IntPtg) t).getValue())),
                new WhatIf(p, lessEqualPtg, t -> leq()),
                new WhatIf(p, lessThanPtg, t -> lt()),
                new WhatIf(p, memErrPtg, (Ptg t) -> parseMemErrPtg((MemErrPtg) t)),
                new WhatIf(p, missingArgPtg, (Ptg t) -> parseMissingArgPtg(row, column)),
                new WhatIf(p, multiplyPtg, t -> mult()),
                new WhatIf(p, namePtg, (Ptg t) -> parseNamePtg((NamePtg) t)),
                new WhatIf(p, notEqualPtg, t -> neq()),
                new WhatIf(p, numberPtg, t -> FLOAT(((NumberPtg) t).getValue())),
                new WhatIf(p, parenthesisPtg, t -> ParenthesisFormula()),
                new WhatIf(p, percentPtg, t -> percentFormula()),
                new WhatIf(p, powerPtg, t -> power()),
                new WhatIf(p, ref3DPxg, (Ptg t) -> parseRef3DPxg((Ref3DPxg) t)),
                new WhatIf(p, refErrorPtg, (Ptg t) -> parseRefErrorPtg((RefErrorPtg) t)),
                new WhatIf(p, refPtg, (Ptg t) -> parseRefPtg((RefPtg) t)),
                new WhatIf(p, stringPtg, (Ptg t) -> TEXT(((StringPtg) t).getValue())),
                new WhatIf(p, subtractPtg, t -> sub()),
                new WhatIf(p, unaryMinusPtg, (Ptg t) -> Minus()),
                new WhatIf(p, unaryPlusPtg, (Ptg t) -> Plus()),
                new WhatIf(p, unionPtg, t -> union((UnionPtg) t)),
                new WhatIf(p, unknownPtg, (Ptg t) -> parseUnknownPtg((UnknownPtg) t))
        )) {
            stream.filter((WhatIf t) -> t.predicate.test(t.ptg)).forEach(t -> t.consumer.accept(t.ptg));
        } catch (Exception e) {
            System.err.println("parse: " + p.getClass().getSimpleName());
            System.err.println(this.currentSheetName + "row:" + row + "column:" + column + e.getMessage());
            e.printStackTrace();
            //System.exit(-1);
        }
    }


    private void parseArrayPtg(ArrayPtg t) {
        Object[][] tokens = t.getTokenArrayValues();
        ConstantArray(tokens);
    }

    /**
     * Area3DPxg is XSSF Area 3D Reference (Sheet + Area) Defined an area in an
     * external or different sheet.
     * <p>
     * This is XSSF only, as it stores the sheet / book references in String
     * form. The HSSF equivalent using indexes is Area3DPtg
     *
     * @param t
     */
    private void parseArea3DPxg(Area3DPxg t) {
        String sheetName = t.getSheetName();
        int sheetIndex = evaluationWorkbook.getSheetIndex(sheetName);
        SHEET tSHEET = new SHEET(sheetName, sheetIndex);
        String area = t.format2DRefAsString();
        RangeInternal range = new RangeInternal(workbook, t.getSheetName(), t);
        parseArea3D(range.getRANGE(), range.getFirstRow(), range.getFirstColumn(), range.getLastRow(), range.getLastColumn(),
                range.getRANGE().values(), tSHEET, area);


    }

    /**
     * Title: XSSF 3D Reference
     * <p>
     * Description: Defines a cell in an external or different sheet.
     * <p>
     * REFERENCE:
     * This is XSSF only, as it stores the sheet / book references in String form. The HSSF equivalent using indexes is Ref3DPtg
     *
     * @param t
     */
    private void parseRef3DPxg(Ref3DPxg t) {
        int extWorkbookNumber = t.getExternalWorkbookNumber();
        String sheet_ = t.getSheetName();
        int sheetIndex = evaluationWorkbook.getSheetIndex(sheet_);
        SHEET tSHEET = new SHEET(sheet_, sheetIndex);
        FILE tFILE = new FILE(extWorkbookNumber, tSHEET);
        String cellref = t.format2DRefAsString();
        if (extWorkbookNumber > 0) parseReference(tFILE, cellref);
        else parseReference(tSHEET, cellref);
    }

    private void parseAttrPtg(AttrPtg t) {
        if (t.isSum()) sum();
    }

    private void parseAreaPtg(AreaPtg t) {
        RangeInternal range = new RangeInternal(workbook, sheet, t);
        rangeReference(range.getRANGE());
    }

    private void parseErrPtg(ErrPtg t) {
        ErrInternal err = new ErrInternal(t);
        ERROR(err.text());
    }

    private void parseFuncPtg(FuncPtg t) {
        parseFunc(t.getName(), t.getNumberOfOperands(), t.isExternalFunction());
    }

    private void parseFuncVarPtg(FuncVarPtg t) {
        parseFunc(t.getName(), t.getNumberOfOperands(), t.isExternalFunction());
    }

    private void parseNamePtg(NamePtg t) {
        EvaluationName evaluationName = evaluationWorkbook.getName(t);
        RangeInternal range = null;
        Ptg[] ptgs = evaluationName.getNameDefinition();
        for (Ptg ptg : ptgs) {
            if (ptg != null) {
                if (ptg instanceof Area3DPxg) {
                    Area3DPxg area3DPxg = (Area3DPxg) ptg;
                    range = new RangeInternal(workbook, area3DPxg.getSheetName(), area3DPxg);
                }
            }
        }

        String name = evaluationWorkbook.getNameText(t);
        namedRange(Objects.requireNonNull(range).getFirstRow(), range.getFirstColumn(), range.getLastRow(), range.getLastColumn(), range.getRANGE().values(), name, range.getSheetName());
    }


    private void parseRefPtg(RefPtg t) {
        int ri = t.getRow();
        int ci = t.getColumn();
        Row row = sheet.getRow(t.getRow());
        boolean rowNotNull = (row != null);
        Object value = null;
        String comment = null;
        if (rowNotNull) {
            Cell c = row.getCell(ci);
            CellInternal excelType = new CellInternal(c);
            value = excelType.valueOf();
            comment = excelType.getComment();
        }
        CELL_REFERENCE(ri, ci, rowNotNull, value, comment);
    }

    private void parseRefErrorPtg(RefErrorPtg t) {
        String text = t.toString();
        ERROR_REF(text);
    }

    private void parseMemErrPtg(MemErrPtg t) {
        err("MemErrPtg: " + t.toString(), formulaRow, formulaColumn);
    }

    private void parseDeleted3DPxg(Deleted3DPxg t) {
        err("Deleted3DPxg: " + t.toString(), formulaRow, formulaColumn);
    }

    private void parseDeletedRef3DPtg(DeletedRef3DPtg t) {
        err("DeletedRef3DPtg: " + t.toString(), formulaRow, formulaColumn);
    }

    private void parseMissingArgPtg(int row, int column) {
        parseMissingArguments(row, column);
    }

    private void parseDeletedArea3DPtg(DeletedArea3DPtg t) {
        err("DeletedArea3DPtg: " + t.toString(), formulaRow, formulaColumn);
    }

    private void parseAreaErrPtg(AreaErrPtg t) {
        err("AreaErrPtg: " + t.toString(), formulaRow, formulaColumn);
    }

    private void parseUnknownPtg(UnknownPtg t) {
        err("Error Unknown Ptg: " + t.toString(), formulaRow, formulaColumn);
    }

    protected abstract void parseFormula(Start start);

    protected abstract void parseMissingArguments(int row, int column);

    protected abstract void ConstantArray(Object[][] array);

    protected abstract void add();

    protected abstract void parseArea3D(RANGE tRANGE, int FirstRow, int FirstColumn, int LastRow, int LastColumn, List<Object> list, SHEET tSHEET, String area);

    protected abstract void sum();

    protected abstract void rangeReference(RANGE tRANGE);

    protected abstract void BOOL(Boolean bool);

    protected abstract void concat();

    protected abstract void div();

    protected abstract void eq();

    protected abstract void ERROR(String text);

    protected abstract void parseFunc(String name, int arity, boolean externalFunction);

    protected abstract void gteq();

    protected abstract void gt();

    protected abstract void intersection();

    protected abstract void INT(Integer value);

    protected abstract void leq();

    protected abstract void lt();

    protected abstract void mult();

    protected abstract void namedRange(int firstRow, int firstColumn, int lastRow, int lastColumn, List<Object> cells, String name, String sheetName);

    protected abstract void neq();

    protected abstract void FLOAT(Double value);

    protected abstract void ParenthesisFormula();

    protected abstract void percentFormula();

    protected abstract void power();

    protected abstract void parseReference(FILE tFILE, String area);

    protected abstract void parseReference(SHEET tSHEET, String area);

    protected abstract void ERROR_REF(String text);

    protected abstract void CELL_REFERENCE(int ri, int ci, boolean rowNotNull, Object value, String comment);

    protected abstract void TEXT(String string);

    protected abstract void sub();

    protected abstract void Minus();

    protected abstract void Plus();

    private void union(UnionPtg t) {
        union();
    }

    protected abstract void union();

    protected abstract void doesFormulaReferToDeletedCell(int row, int column);

    protected abstract void parseFormulaInit();

    protected abstract Start parseFormulaPost();

    static final class ErrInternal {

        private final static String ERROR_NULL_INTERSECTION = "#NULL!";
        private final static String ERROR_DIV_ZERO = "#DIV/0!";
        private final static String ERROR_VALUE_INVALID = "#VALUE!";
        private final static String ERROR_REF_INVALID = "#REF!";
        private final static String ERROR_NAME_INVALID = "#NAME?";
        private final static String ERROR_NUM_ERROR = "#NUM!";
        private final static String ERROR_N_A = "#N/A";

        private final ErrPtg t;

        ErrInternal(ErrPtg t) {
            this.t = t;
        }

        String text() {
            if (t == NULL_INTERSECTION) return ERROR_NULL_INTERSECTION;
            else if (t == DIV_ZERO) return ERROR_DIV_ZERO;
            else if (t == VALUE_INVALID) return ERROR_VALUE_INVALID;
            else if (t == REF_INVALID) return ERROR_REF_INVALID;
            else if (t == NAME_INVALID) return ERROR_NAME_INVALID;
            else if (t == NUM_ERROR) return ERROR_NUM_ERROR;
            else if (t == N_A) return ERROR_N_A;
            else return "FIXME!";
        }
    }

    // 3DPxg is XSSF
    // 3DPtg is HSSF
    class WhatIf {

        final Ptg ptg;
        final Predicate<Ptg> predicate;
        final Consumer<Ptg> consumer;

        WhatIf(Ptg ptg, Predicate<Ptg> predicate, Consumer<Ptg> consumer) {
            this.ptg = ptg;
            this.predicate = predicate;
            this.consumer = consumer;
        }
    }

    /**
     * @author mcaliman
     */
    class RangeInternal {

        private final SpreadsheetVersion SPREADSHEET_VERSION = SpreadsheetVersion.EXCEL2007;
        private final Workbook workbook;
        private final Sheet sheet;

        private final int firstRow;
        private final int firstColumn;

        private final int lastRow;
        private final int lastColumn;
        RANGE tRANGE;
        private List<Object> values;
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

            values = new ArrayList<>();
            for (Cell cell : cells)
                if (cell != null) {
                    CellInternal excelType = new CellInternal(cell);
                    values.add(excelType.valueOf());
                    tRANGE.add(excelType.valueOf());
                }
        }

        private void init() {
            String refs = tRANGE.toString();
            List<Cell> cells = range(refs);
            values = new ArrayList<>();
            for (Cell cell : cells)
                if (cell != null) {
                    CellInternal excelType = new CellInternal(cell);
                    values.add(excelType.valueOf());
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

        int getFirstRow() {
            return firstRow;
        }

        int getFirstColumn() {
            return firstColumn;
        }

        int getLastRow() {
            return lastRow;
        }

        int getLastColumn() {
            return lastColumn;
        }

        String getSheetName() {
            return sheetName;
        }

        RANGE getRANGE() {
            return tRANGE;
        }
    }

    /**
     * @author mcaliman
     */
    class CellInternal {

        private final Cell cell;
        private final String comment;

        @SuppressWarnings("unused")
        CellInternal(Cell cell) {
            this.cell = cell;
            Comment cellComment = this.cell.getCellComment();
            comment = comment(cellComment);
            CellStyle style = this.cell.getCellStyle();
            String format = style.getDataFormatString();
        }

        String getComment() {
            return comment;
        }

        private String comment(Comment comment) {
            if (comment == null) return null;
            RichTextString text = comment.getString();
            if (text == null) return null;
            return text.getString();

        }

        private Object valueOf() {
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

        private boolean isDataType(Cell c) {
            return c.getCellType() == CELL_TYPE_NUMERIC && HSSFDateUtil.isCellDateFormatted(c);
        }

        Class internalFormulaResultType() {
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

    class FormulaTokensInternal {

        private final XSSFEvaluationWorkbook ew;
        private final EvaluationSheet es;

        FormulaTokensInternal(XSSFEvaluationWorkbook ew, EvaluationSheet es) {
            this.ew = ew;
            this.es = es;
        }

        Ptg[] getFormulaTokens(int row, int column) {
            EvaluationCell evalCell = es.getCell(row, column);
            Ptg[] ptgs = null;
            try {
                ptgs = ew.getFormulaTokens(evalCell);
            } catch (FormulaParseException e) {
                err("" + e.getMessage(), row, column);
            }
            return ptgs;
        }

        private void err(String string, int row, int column) {
            System.err.println(string + " row:" + row + " col:" + column);
        }
    }
}
