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

import excel.grammar.Start;
import excel.grammar.formula.reference.*;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.ptg.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;

import java.io.File;
import java.io.IOException;
import java.util.Objects;
import java.util.function.Predicate;
import java.util.stream.Stream;

import static org.apache.poi.ss.formula.ptg.ErrPtg.*;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA;

/**
 * 3DPxg is XSSF
 * 3DPtg is HSSF
 *
 * @author Massimo Caliman
 */
public abstract class AbstractParser {

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

    /**
     * (Work)Book
     */
    private final Workbook book;
    /**
     * (Work)Sheet
     */
    private Sheet sheet;

    public boolean verbose = false;
    public boolean metadata = false;
    protected boolean errors = false;
    /**
     * Current Formula Colum
     */
    protected int colFormula;
    /**
     * Current Formula Row
     */
    protected int rowFormula;
    /**
     * Current Sheet Index
     */
    protected int sheetIndex;
    /**
     * Current Sheet Name
     */
    protected String sheetName;

    /**
     * (Work)Book Protection Present flag
     */
    private boolean protectionPresent;
    private String fileName;

    private final Helper helper;

    protected AbstractParser(File file) throws InvalidFormatException, IOException {
        this(WorkbookFactory.create(file));
        this.fileName = file.getName();
    }

    private AbstractParser(Workbook workbook) {
        this.book = workbook;
        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) this.book;
        protectionPresent = xssfWorkbook.validateWorkbookPassword("password");
        POIXMLProperties props = xssfWorkbook.getProperties();
        POIXMLProperties.CoreProperties coreProperties = props.getCoreProperties();
        String creator = coreProperties.getCreator();
        String description = coreProperties.getDescription();
        String keywords = coreProperties.getKeywords();
        String title = coreProperties.getTitle();
        String subject = coreProperties.getSubject();
        String category = coreProperties.getCategory();
        POIXMLProperties.CustomProperties customProperties = props.getCustomProperties();
        POIXMLProperties.ExtendedProperties extendedProperties = props.getExtendedProperties();
        String company = extendedProperties.getUnderlyingProperties().getCompany();
        String template = extendedProperties.getUnderlyingProperties().getTemplate();
        String manager = extendedProperties.getUnderlyingProperties().getManager();
        this.helper = new Helper(this.book);
        System.out.println("Parse...");
    }

    public String getFileName() {
        return fileName;
    }

    public boolean isProtectionPresent() {
        return protectionPresent;
    }

    protected void verbose(String text) {
        if (this.verbose) System.out.println(text);
    }

    protected void err(String string, int row, int column) {

    }

    /**
     * Parse (Work)Book.
     */
    public void parse() {
        for (Sheet currentSheet : this.book) parse(currentSheet);
    }

    /**
     * Parse a single (Work)Sheet
     *
     * @param sheet
     */
    private void parse(@NotNull Sheet sheet) {
        this.sheet = sheet;
        initSheetData();
        for (Row row : sheet)
            for (Cell cell : row)
                if (cell != null) parse(cell);
                else err("Cell is null.", rowFormula, colFormula);
    }

    /**
     * Parse Cell
     *
     * @param cell
     */
    private void parse(Cell cell) {
        if (cell.getCellType() == CELL_TYPE_FORMULA)
            parseFormula(cell);
    }

    /**
     * Parse Formula Cell
     *
     * @param cell
     */
    private void parseFormula(Cell cell) {
        verbose("Cell:" + cell.getClass().getSimpleName() + " " + cell.toString() + " " + cell.getCellType());
        CellInternal excelCell = new CellInternal(cell);
        String comment = excelCell.getComment();
        colFormula = cell.getColumnIndex();
        rowFormula = cell.getRowIndex();
        //Class internalFormulaResultTypeClass = excelCell.internalFormulaResultType();
        String formulaAddress = Start.cellAddress(rowFormula, colFormula);
        String formulaText = cell.getCellFormula();
        verbose(formulaAddress + " = " + formulaText);
        Ptg[] formulaPtgs = helper.tokens(this.sheet,this.rowFormula,this.colFormula);
        if (formulaPtgs == null) {
            System.err.println("ptgs empty or null for address " + formulaAddress);
            err("ptgs empty or null for address " + formulaAddress, rowFormula, colFormula);
            parseUDF(formulaText);
            return;
        }
        Start start = parse(formulaPtgs);
        if (Objects.nonNull(start)) {
            start.setComment(comment);
            parseFormula(start);
        }
    }

    /**
     * Parse Ptg array
     *
     * @param ptgs
     * @return
     */
    private Start parse(Ptg[] ptgs) {
        parseFormulaInit();
        if (Ptg.doesFormulaReferToDeletedCell(ptgs)) doesFormulaReferToDeletedCell(rowFormula, colFormula);
        for (Ptg ptg : ptgs) parse(ptg, rowFormula, colFormula);
        return parseFormulaPost();
    }

    private void parse(@NotNull Ptg p, int row, int column) {
        verbose("parse: " + p.getClass().getSimpleName());
        try (Stream<WhatIf> stream = Stream.of(
                new WhatIf(p, arrayPtg, (Ptg t) -> parseArrayPtg((ArrayPtg) t)),
                new WhatIf(p, addPtg, (Ptg t) -> parseAdd()),
                new WhatIf(p, area3DPxg, (Ptg t) -> parseArea3DPxg((Area3DPxg) t)),
                new WhatIf(p, areaErrPtg, (Ptg t) -> parseAreaErrPtg((AreaErrPtg) t)),
                new WhatIf(p, areaPtg, (Ptg t) -> parseAreaPtg((AreaPtg) t)),
                new WhatIf(p, attrPtg, (Ptg t) -> parseAttrPtg((AttrPtg) t)),
                new WhatIf(p, boolPtg, t -> parseBOOL(((BoolPtg) t).getValue())),
                new WhatIf(p, concatPtg, t -> parseConcat()),
                new WhatIf(p, deleted3DPxg, (Ptg t) -> parseDeleted3DPxg((Deleted3DPxg) t)),
                new WhatIf(p, deletedArea3DPtg, (Ptg t) -> parseDeletedArea3DPtg((DeletedArea3DPtg) t)),
                new WhatIf(p, deletedRef3DPtg, (Ptg t) -> parseDeletedRef3DPtg((DeletedRef3DPtg) t)),
                new WhatIf(p, dividePtg, t -> parseDiv()),
                new WhatIf(p, equalPtg, t -> parseEq()),
                new WhatIf(p, errPtg, (Ptg t) -> parseErrPtg((ErrPtg) t)),
                new WhatIf(p, funcPtg, (Ptg t) -> parseFuncPtg((FuncPtg) t)),
                new WhatIf(p, funcVarPtg, (Ptg t) -> parseFuncVarPtg((FuncVarPtg) t)),
                new WhatIf(p, greaterEqualPtg, t -> parseGteq()),
                new WhatIf(p, greaterThanPtg, t -> parseGt()),
                new WhatIf(p, intersectionPtg, t -> parseIntersection()),
                new WhatIf(p, intPtg, t -> parseINT(((IntPtg) t).getValue())),
                new WhatIf(p, lessEqualPtg, t -> parseLeq()),
                new WhatIf(p, lessThanPtg, t -> parseLt()),
                new WhatIf(p, memErrPtg, (Ptg t) -> parseMemErrPtg((MemErrPtg) t)),
                new WhatIf(p, missingArgPtg, (Ptg t) -> parseMissingArgPtg(row, column)),
                new WhatIf(p, multiplyPtg, t -> parseMult()),
                new WhatIf(p, namePtg, (Ptg t) -> parseNamePtg((NamePtg) t)),
                new WhatIf(p, notEqualPtg, t -> parseNeq()),
                new WhatIf(p, numberPtg, t -> parseFLOAT(((NumberPtg) t).getValue())),
                new WhatIf(p, parenthesisPtg, t -> parseParenthesisFormula()),
                new WhatIf(p, percentPtg, t -> percentFormula()),
                new WhatIf(p, powerPtg, t -> parsePower()),
                new WhatIf(p, ref3DPxg, (Ptg t) -> parseRef3DPxg((Ref3DPxg) t)),
                new WhatIf(p, refErrorPtg, (Ptg t) -> parseRefErrorPtg((RefErrorPtg) t)),
                new WhatIf(p, refPtg, (Ptg t) -> parseRefPtg((RefPtg) t)),
                new WhatIf(p, stringPtg, (Ptg t) -> parseTEXT(((StringPtg) t).getValue())),
                new WhatIf(p, subtractPtg, t -> parseSub()),
                new WhatIf(p, unaryMinusPtg, (Ptg t) -> parseMinus()),
                new WhatIf(p, unaryPlusPtg, (Ptg t) -> parsePlus()),
                new WhatIf(p, unionPtg, t -> parseUnion((UnionPtg) t)),
                new WhatIf(p, unknownPtg, (Ptg t) -> parseUnknownPtg((UnknownPtg) t))
        )) {
            stream.filter((WhatIf t) -> t.predicate.test(t.ptg)).forEach(t -> t.consumer.accept(t.ptg));
        } catch (Exception e) {
            System.err.println("parse: " + p.getClass().getSimpleName());
            System.err.println(this.sheetName + "row:" + row + "column:" + column + e.getMessage());
            e.printStackTrace();
            //System.exit(-1);
        }
    }

    private void initSheetData() {
        protectionPresent = protectionPresent || ((XSSFSheet) sheet).validateSheetPassword("password");
        sheetIndex = book.getSheetIndex(sheet);
        sheetName = sheet.getSheetName();
    }

    private void parseArrayPtg(@NotNull ArrayPtg t) {
        Object[][] tokens = t.getTokenArrayValues();
        parseConstantArray(tokens);
    }

    protected abstract void parseConstantArray(Object[][] array);

    protected abstract void parseUDF(String arguments);



    /**
     * Area3DPxg is XSSF Area 3D Reference (Sheet + Area) Defined an area in an
     * external or different sheet.
     * <p>
     * This is XSSF only, as it stores the sheet / book references in String
     * form. The HSSF equivalent using indexes is Area3DPtg
     *
     * @param t
     */
    private void parseArea3DPxg(@NotNull Area3DPxg t) {
        String sheetName = t.getSheetName();
        int sheetIndex = helper.getSheetIndex(sheetName);
        SHEET tSHEET = new SHEET(sheetName, sheetIndex);

        String area = t.format2DRefAsString();
        RangeInternal range = new RangeInternal(book, t.getSheetName(), t);
        parseArea3D(range.getRANGE(), tSHEET, area);
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
    private void parseRef3DPxg(@NotNull Ref3DPxg t) {
        int extWorkbookNumber = t.getExternalWorkbookNumber();

        String sheet_ = t.getSheetName();
        int sheetIndex = helper.getSheetIndex(sheet_);

        SHEET tSHEET = new SHEET(sheet_, sheetIndex);
        FILE tFILE = new FILE(extWorkbookNumber, tSHEET);
        String cellref = t.format2DRefAsString();
        if (extWorkbookNumber > 0) parseReference(tFILE, cellref);
        else parseReference(tSHEET, cellref);
    }

    private void parseAttrPtg(@NotNull AttrPtg t) {
        if (t.isSum()) parseSum();
    }

    private void parseAreaPtg(AreaPtg t) {
        RangeInternal range = new RangeInternal(book, sheet, t);
        parseRangeReference(range.getRANGE());
    }

    private void parseErrPtg(ErrPtg t) {
        ErrInternal err = new ErrInternal(t);
        parseERROR(err.text());
    }

    private void parseFuncPtg(@NotNull FuncPtg t) {
        if (t.getNumberOfOperands() == 0) parseFunc(t.getName(), t.isExternalFunction());
        else parseFunc(t.getName(), t.getNumberOfOperands(), t.isExternalFunction());
    }

    private void parseFuncVarPtg(@NotNull FuncVarPtg t) {
        if (t.getNumberOfOperands() == 0) parseFunc(t.getName(), t.isExternalFunction());
        else parseFunc(t.getName(), t.getNumberOfOperands(), t.isExternalFunction());
    }

    private void parseNamePtg(NamePtg t) {
        RangeInternal range = null;

        Ptg[] ptgs = helper.getName(t);
        String name = helper.getNameText(t);

        for (Ptg ptg : ptgs) {
            if (ptg != null) {
                if (ptg instanceof Area3DPxg) {
                    Area3DPxg area3DPxg = (Area3DPxg) ptg;
                    range = new RangeInternal(book, area3DPxg.getSheetName(), area3DPxg);
                }
            }
        }

        RANGE tRANGE = range.getRANGE();
        parseNamedRange(tRANGE, name, range.getSheetName());
    }

    private void parseRefPtg(@NotNull RefPtg t) {
        Row rowObject = sheet.getRow(t.getRow());
        Object value = null;
        String comment = null;
        if (rowObject != null) {
            Cell c = rowObject.getCell(t.getColumn());
            CellInternal excelType = new CellInternal(c);
            value = excelType.valueOf();
            comment = excelType.getComment();
        }
        CELL_REFERENCE tCELL_REFERENCE = new CELL_REFERENCE(t.getRow(), t.getColumn(), comment);

        parseCELL_REFERENCE(tCELL_REFERENCE, rowObject != null, value);
    }

    private void parseRefErrorPtg(RefErrorPtg t) {
        ERROR_REF term = new ERROR_REF();
        parseERROR_REF(term);
    }

    protected abstract void parseERROR_REF(ERROR_REF ref);

    private void parseMemErrPtg(@NotNull MemErrPtg t) {
        err("MemErrPtg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseDeleted3DPxg(@NotNull Deleted3DPxg t) {
        err("Deleted3DPxg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseDeletedRef3DPtg(@NotNull DeletedRef3DPtg t) {
        err("DeletedRef3DPtg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseMissingArgPtg(int row, int column) {
        parseMissingArguments(row, column);
    }

    private void parseDeletedArea3DPtg(@NotNull DeletedArea3DPtg t) {
        err("DeletedArea3DPtg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseAreaErrPtg(@NotNull AreaErrPtg t) {
        err("AreaErrPtg: " + t.toString(), rowFormula, colFormula);
    }

    private void parseUnknownPtg(@NotNull UnknownPtg t) {
        err("Error Unknown Ptg: " + t.toString(), rowFormula, colFormula);
    }

    protected abstract void parseFormula(Start start);

    protected abstract void parseMissingArguments(int row, int column);

    protected abstract void parseAdd();

    protected abstract void parseArea3D(RANGE tRANGE, SHEET tSHEET, String area);

    protected abstract void parseSum();

    protected abstract void parseRangeReference(RANGE tRANGE);

    protected abstract void parseBOOL(Boolean bool);

    protected abstract void parseConcat();

    protected abstract void parseDiv();

    protected abstract void parseEq();

    protected abstract void parseERROR(String text);

    protected abstract void parseFunc(String name, int arity, boolean externalFunction);

    protected abstract void parseFunc(String name, boolean externalFunction);

    protected abstract void parseGteq();

    protected abstract void parseGt();

    protected abstract void parseIntersection();

    protected abstract void parseINT(Integer value);

    protected abstract void parseLeq();

    protected abstract void parseLt();

    protected abstract void parseMult();

    protected abstract void parseNamedRange(RANGE tRANGE, String name, String sheetName);

    protected abstract void parseNeq();

    protected abstract void parseFLOAT(Double value);

    protected abstract void parseParenthesisFormula();

    protected abstract void percentFormula();

    protected abstract void parsePower();

    protected abstract void parseReference(FILE tFILE, String area);

    protected abstract void parseReference(SHEET tSHEET, String area);



    protected abstract void parseCELL_REFERENCE(CELL_REFERENCE tCELL_REFERENCE, boolean rowNotNull, Object value);

    protected abstract void parseTEXT(String string);

    protected abstract void parseSub();

    protected abstract void parseMinus();

    protected abstract void parsePlus();

    private void parseUnion(UnionPtg t) {
        parseUnion();
    }

    protected abstract void parseUnion();

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




}
