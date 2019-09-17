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

import dev.caliman.excel.grammar.Start;
import dev.caliman.excel.grammar.formula.constant.*;
import dev.caliman.excel.grammar.formula.reference.*;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.ptg.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
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

    protected final boolean errors = false;
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
    @NotNull
    private final Helper helper;
    private final List<Cell> ext;
    public boolean verbose = false;
    public boolean metadata = false;
    /**
     * Current Formula Column
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
    protected boolean isSingleSheet;

    //Meta
    private String creator;
    private String description;
    private String keywords;
    private String title;
    private String subject;
    private String category;

    private int counterSheets = 0;
    private int counterFormulas;
    /**
     * (Work)Sheet
     */
    private Sheet sheet;
    /**
     * (Work)Book Protection Present flag
     */
    private boolean protectionPresent;
    private String fileName;


    protected AbstractParser(@NotNull File file) throws InvalidFormatException, IOException {
        this(WorkbookFactory.create(file));
        this.fileName = file.getName();
    }

    private AbstractParser(Workbook workbook) {
        this.book = workbook;
        this.ext = new ArrayList<>();
        this.helper = new Helper(this.book);
        readMetadata();
        print();
    }

    private void readMetadata() {
        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) this.book;
        this.protectionPresent = xssfWorkbook.validateWorkbookPassword("password");
        POIXMLProperties props = xssfWorkbook.getProperties();
        POIXMLProperties.CoreProperties coreProperties = props.getCoreProperties();

        this.creator = coreProperties.getCreator();
        this.description = coreProperties.getDescription();
        this.keywords = coreProperties.getKeywords();
        this.title = coreProperties.getTitle();
        this.subject = coreProperties.getSubject();
        this.category = coreProperties.getCategory();
    }

    public int getCounterFormulas() {
        return counterFormulas;
    }

    public String getFileName() {
        return fileName;
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
        this.isSingleSheet = this.book.getNumberOfSheets() == 1;
        for (Sheet currentSheet : this.book) parse(currentSheet);
    }

    /**
     * Parse a single (Work)Sheet
     */
    private void parse(@NotNull Sheet sheet) {
        this.counterSheets++;
        this.sheet = sheet;
        protectionPresent = protectionPresent || ((XSSFSheet) sheet).validateSheetPassword("password");
        this.sheetIndex = book.getSheetIndex(sheet);
        this.sheetName = sheet.getSheetName();
        verbose("Parsing sheet-name:" + this.sheetName);
        for (Row row : sheet)
            for (Cell cell : row)
                if (cell != null) parse(cell);
                else err("Cell is null.", rowFormula, colFormula);
    }

    /**
     * Parse Cell
     */
    private void parse(Cell cell) {
        if (cell.getCellType() == CELL_TYPE_FORMULA) {
            parseFormula(cell);
            this.counterFormulas++;
        } else if (this.ext.contains(cell)) {
            verbose("Recover loosed cell!");
            Object obj = Helper.valueOf(cell);
            CELL_REFERENCE cell_reference = new CELL_REFERENCE(cell.getRowIndex(), cell.getColumnIndex());
            cell_reference.setValue(obj);
            cell_reference.setSheetName(cell.getSheet().getSheetName());
            cell_reference.setSheetIndex(helper.getSheetIndex(cell.getSheet().getSheetName()));
            parseCELL_REFERENCELinked(cell_reference);
            this.ext.remove(cell);
        }
    }

    /**
     * Parse Formula Cell
     */
    private void parseFormula(Cell cell) {
        verbose("Cell:" + cell.getClass().getSimpleName() + " " + cell.toString() + " " + cell.getCellType());
        String comment = Helper.getComment(cell);
        colFormula = cell.getColumnIndex();
        rowFormula = cell.getRowIndex();
        String formulaAddress = Start.cellAddress(rowFormula, colFormula);
        //String formulaText = cell.getCellFormula();
        //verbose(formulaAddress + " = " + formulaText);
        Ptg[] formulaPtgs = helper.tokens(this.sheet, this.rowFormula, this.colFormula);
        if (formulaPtgs == null) {
            String formulaText = cell.getCellFormula();
            System.err.println("ptgs empty or null for address " + formulaAddress);
            err("ptgs empty or null for address " + formulaAddress, rowFormula, colFormula);
            parseUDF(formulaText);
            return;
        }
        Start start = parse(formulaPtgs);
        if (Objects.nonNull(start)) {
            start.setComment(comment);
            start.setSingleSheet(this.isSingleSheet);
            parseFormula(start);
        }
    }

    /**
     * Parse Ptg array
     *
     * @param ptgs
     * @return
     */
    @SuppressWarnings("JavaDoc")
    private Start parse(@NotNull Ptg[] ptgs) {
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
                new WhatIf(p, boolPtg, t -> parseBooleanLiteral(((BoolPtg) t).getValue())),
                new WhatIf(p, concatPtg, t -> parseConcat()),
                new WhatIf(p, deleted3DPxg, (Ptg t) -> parseDeleted3DPxg((Deleted3DPxg) t)),
                new WhatIf(p, deletedArea3DPtg, (Ptg t) -> parseDeletedArea3DPtg((DeletedArea3DPtg) t)),
                new WhatIf(p, deletedRef3DPtg, (Ptg t) -> parseDeletedRef3DPtg((DeletedRef3DPtg) t)),
                new WhatIf(p, dividePtg, t -> parseDiv()),
                new WhatIf(p, equalPtg, t -> parseEq()),
                new WhatIf(p, errPtg, (Ptg t) -> parseErrorLiteral((ErrPtg) t)),
                new WhatIf(p, funcPtg, (Ptg t) -> parseFuncPtg((FuncPtg) t)),
                new WhatIf(p, funcVarPtg, (Ptg t) -> parseFuncVarPtg((FuncVarPtg) t)),
                new WhatIf(p, greaterEqualPtg, t -> parseGteq()),
                new WhatIf(p, greaterThanPtg, t -> parseGt()),
                new WhatIf(p, intersectionPtg, t -> parseIntersection()),
                new WhatIf(p, intPtg, t -> parseIntLiteral(((IntPtg) t).getValue())),
                new WhatIf(p, lessEqualPtg, t -> parseLeq()),
                new WhatIf(p, lessThanPtg, t -> parseLt()),
                new WhatIf(p, memErrPtg, (Ptg t) -> parseMemErrPtg((MemErrPtg) t)),
                new WhatIf(p, missingArgPtg, (Ptg t) -> parseMissingArgPtg(row, column)),
                new WhatIf(p, multiplyPtg, t -> parseMult()),
                new WhatIf(p, namePtg, (Ptg t) -> parseNamePtg((NamePtg) t)),
                new WhatIf(p, notEqualPtg, t -> parseNeq()),
                new WhatIf(p, numberPtg, t -> parseFloatLiteral(((NumberPtg) t).getValue())),
                new WhatIf(p, parenthesisPtg, t -> parseParenthesisFormula()),
                new WhatIf(p, percentPtg, t -> percentFormula()),
                new WhatIf(p, powerPtg, t -> parsePower()),
                new WhatIf(p, ref3DPxg, (Ptg t) -> parseRef3DPxg((Ref3DPxg) t)),
                new WhatIf(p, refErrorPtg, (Ptg t) -> parseReferenceErrorLiteral()),
                new WhatIf(p, refPtg, (Ptg t) -> parseRefPtg((RefPtg) t)),
                new WhatIf(p, stringPtg, (Ptg t) -> parseStringLiteral(((StringPtg) t).getValue())),
                new WhatIf(p, subtractPtg, t -> parseSub()),
                new WhatIf(p, unaryMinusPtg, (Ptg t) -> parseMinus()),
                new WhatIf(p, unaryPlusPtg, (Ptg t) -> parsePlus()),
                new WhatIf(p, unionPtg, t -> parseUnion()),
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

    protected abstract void doesFormulaReferToDeletedCell(int row, int column);

    protected abstract void parseFormulaInit();

    protected abstract Start parseFormulaPost();

    //region Reference

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

        String area = helper.getArea(t);
        parseArea3D(helper.getRANGE(sheetName, t), tSHEET, area);
    }

    protected abstract void parseArea3D(RANGE tRANGE, SHEET tSHEET, String area);

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
        String sheetName = t.getSheetName();
        int sheetIndex = helper.getSheetIndex(sheetName);
        SHEET tSHEET = new SHEET(sheetName, sheetIndex);
        FILE tFILE = new FILE(extWorkbookNumber, tSHEET);
        String cellref = helper.getCellRef(t);
        if (this.sheetIndex != sheetIndex) {
            Sheet extSheet = this.book.getSheet(sheetName);
            if (extSheet != null) {
                CellReference cr = new CellReference(cellref);
                Row row = extSheet.getRow(cr.getRow());
                Cell cell = row.getCell(cr.getCol());
                this.ext.add(cell);
                verbose("Loosing!!! reference[ext] " + tSHEET.toString() + "" + cellref);
            }
        }
        if (extWorkbookNumber > 0) parseReference(tFILE, cellref);
        else parseReference(tSHEET, cellref);
    }

    protected abstract void parseReference(FILE tFILE, String area);

    private void parseAreaPtg(@NotNull AreaPtg t) {
        parseRangeReference(helper.getRANGE(sheet, t));
    }

    protected abstract void parseRangeReference(RANGE tRANGE);

    private void parseNamePtg(@NotNull NamePtg t) {
        RangeInternal range = null;
        Ptg[] ptgs = helper.getName(t);
        String name = helper.getNameText(t);
        int sheetIndex = 0;
        for (Ptg ptg : ptgs) {
            if (ptg != null) {
                if (ptg instanceof Area3DPxg) {
                    Area3DPxg area3DPxg = (Area3DPxg) ptg;

                    range = new RangeInternal(book, area3DPxg.getSheetName(), area3DPxg);
                    sheetIndex = helper.getSheetIndex(area3DPxg.getSheetName());
                }
            }
        }
        RANGE tRANGE = Objects.requireNonNull(range).getRANGE();

        NamedRange term = new NamedRange(name, tRANGE);
        term.setSheetIndex(sheetIndex);
        term.setSheetName(range.getSheetName());

        parseNamedRange(term);
    }

    protected abstract void parseNamedRange(NamedRange tNamedRange);

    protected abstract void parseReference(SHEET tSHEET, String area);

    private void parseRefPtg(@NotNull RefPtg t) {
        Row rowObject = sheet.getRow(t.getRow());
        Object value = null;
        String comment = null;
        if (rowObject != null) {
            Cell c = rowObject.getCell(t.getColumn());
            value = Helper.valueOf(c);
            comment = Helper.getComment(c);
        }
        CELL_REFERENCE tCELL_REFERENCE = new CELL_REFERENCE(t.getRow(), t.getColumn(), comment);
        tCELL_REFERENCE.setValue(value);
        parseCELL_REFERENCE(tCELL_REFERENCE);
    }

    protected abstract void parseCELL_REFERENCE(CELL_REFERENCE tCELL_REFERENCE);

    protected abstract void parseCELL_REFERENCELinked(CELL_REFERENCE tCELL_REFERENCE);

    //endregion

    //region Formula
    private void parseArrayPtg(@NotNull ArrayPtg t) {
        parseConstantArray(t.getTokenArrayValues());
    }

    protected abstract void parseConstantArray(Object[][] array);

    protected abstract void parseUDF(String arguments);

    private void parseAttrPtg(@NotNull AttrPtg t) {
        if (t.isSum()) parseSum();
    }

    protected abstract void parseParenthesisFormula();

    private void parseFuncVarPtg(@NotNull FuncVarPtg t) {
        if (t.getNumberOfOperands() == 0) parseFunc(t.getName());
        else parseFunc(t.getName(), t.getNumberOfOperands());
    }

    private void parseFuncPtg(@NotNull FuncPtg t) {
        if (t.getNumberOfOperands() == 0) parseFunc(t.getName());
        else parseFunc(t.getName(), t.getNumberOfOperands());
    }

    protected abstract void parseFunc(String name, int arity);

    protected abstract void parseFunc(String name);

    protected abstract void percentFormula();

    protected abstract void parseSum();
    //endregion

    //region Binary
    protected abstract void parseAdd();

    protected abstract void parseSub();

    protected abstract void parseMult();

    protected abstract void parseDiv();

    protected abstract void parsePower();

    protected abstract void parseEq();

    protected abstract void parseGteq();

    protected abstract void parseGt();

    protected abstract void parseLeq();

    protected abstract void parseLt();

    protected abstract void parseNeq();

    protected abstract void parseConcat();
    //endregion

    //region Unary
    protected abstract void parseMinus();

    protected abstract void parsePlus();
    //endregion

    //region Union & Intersection
    protected abstract void parseUnion();

    protected abstract void parseIntersection();
    //endregion

    //region Constants
    //@todo impl. DATETIME
    private void parseErrorLiteral(ErrPtg t) {
        String text;
        if (t == NULL_INTERSECTION) text = "#NULL!";
        else if (t == DIV_ZERO) text = "#DIV/0!";
        else if (t == VALUE_INVALID) text = "#VALUE!";
        else if (t == REF_INVALID) text = "#REF!";
        else if (t == NAME_INVALID) text = "#NAME?";
        else if (t == NUM_ERROR) text = "#NUM!";
        else if (t == N_A) text = "#N/A";
        else text = "FIXME!";

        var term = new ERROR(text);
        parseErrorLiteral(term);
    }

    protected abstract void parseErrorLiteral(ERROR t);

    private void parseBooleanLiteral(Boolean bool) {
        var term = new BOOL(bool);
        parseBooleanLiteral(term);
    }

    protected abstract void parseBooleanLiteral(BOOL bool);

    private void parseStringLiteral(String string) {
        var term = new TEXT(string);
        parseStringLiteral(term);
    }

    protected abstract void parseStringLiteral(TEXT term);

    private void parseIntLiteral(Integer value) {
        var term = new INT(value);
        parseIntLiteral(term);
    }

    protected abstract void parseIntLiteral(INT value);

    private void parseFloatLiteral(Double value) {
        var term = new FLOAT(value);
        parseFloatLiteral(term);
    }

    protected abstract void parseFloatLiteral(FLOAT term);
    //endregion

    //region Literals
    private void parseReferenceErrorLiteral() {
        ERROR_REF term = new ERROR_REF();
        parseReferenceErrorLiteral(term);
    }

    protected abstract void parseReferenceErrorLiteral(ERROR_REF term);
    //endregion

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

    private void print() {
        System.out.println("Parse...");
    }
}
