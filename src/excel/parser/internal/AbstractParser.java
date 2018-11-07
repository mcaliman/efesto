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
import org.apache.poi.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.EvaluationName;
import org.apache.poi.ss.formula.EvaluationSheet;
import org.apache.poi.ss.formula.ptg.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Predicate;
import java.util.stream.Stream;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA;

public abstract class AbstractParser {

    private final Predicate<Ptg> arrayPtg = (Ptg t) -> t instanceof ArrayPtg;
    private final Predicate<Ptg> addPtg = (Ptg t) -> t instanceof AddPtg;
    private final Predicate<Ptg> area3DPxg = (Ptg t) -> t instanceof Area3DPxg;
    private final Predicate<Ptg> areaErrPtg = (Ptg t) -> t instanceof AreaErrPtg;
    private final Predicate<Ptg> areaNPtg = (Ptg t) -> t instanceof AreaNPtg;
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
    private final Predicate<Ptg> expPtg = (Ptg t) -> t instanceof ExpPtg;
    private final Predicate<Ptg> funcPtg = (Ptg t) -> t instanceof FuncPtg;
    private final Predicate<Ptg> funcVarPtg = (Ptg t) -> t instanceof FuncVarPtg;
    private final Predicate<Ptg> greaterEqualPtg = (Ptg t) -> t instanceof GreaterEqualPtg;
    private final Predicate<Ptg> greaterThanPtg = (Ptg t) -> t instanceof GreaterThanPtg;
    private final Predicate<Ptg> intersectionPtg = (Ptg t) -> t instanceof IntersectionPtg;
    private final Predicate<Ptg> intPtg = (Ptg t) -> t instanceof IntPtg;
    private final Predicate<Ptg> lessEqualPtg = (Ptg t) -> t instanceof LessEqualPtg;
    private final Predicate<Ptg> lessThanPtg = (Ptg t) -> t instanceof LessThanPtg;
    private final Predicate<Ptg> memAreaPtg = (Ptg t) -> t instanceof MemAreaPtg;
    private final Predicate<Ptg> memErrPtg = (Ptg t) -> t instanceof MemErrPtg;
    private final Predicate<Ptg> memFuncPtg = (Ptg t) -> t instanceof MemFuncPtg;
    private final Predicate<Ptg> missingArgPtg = (Ptg t) -> t instanceof MissingArgPtg;
    private final Predicate<Ptg> multiplyPtg = (Ptg t) -> t instanceof MultiplyPtg;
    private final Predicate<Ptg> namePtg = (Ptg t) -> t instanceof NamePtg;
    private final Predicate<Ptg> nameXPxg = (Ptg t) -> t instanceof NameXPxg;
    private final Predicate<Ptg> notEqualPtg = (Ptg t) -> t instanceof NotEqualPtg;
    private final Predicate<Ptg> numberPtg = (Ptg t) -> t instanceof NumberPtg;
    private final Predicate<Ptg> parenthesisPtg = (Ptg t) -> t instanceof ParenthesisPtg;
    private final Predicate<Ptg> percentPtg = (Ptg t) -> t instanceof PercentPtg;
    private final Predicate<Ptg> powerPtg = (Ptg t) -> t instanceof PowerPtg;
    private final Predicate<Ptg> rangePtg = (Ptg t) -> t instanceof RangePtg;
    private final Predicate<Ptg> ref3DPxg = (Ptg t) -> t instanceof Ref3DPxg;
    private final Predicate<Ptg> refErrorPtg = (Ptg t) -> t instanceof RefErrorPtg;
    private final Predicate<Ptg> refNPtg = (Ptg t) -> t instanceof RefNPtg;
    private final Predicate<Ptg> refPtg = (Ptg t) -> t instanceof RefPtg;
    private final Predicate<Ptg> stringPtg = (Ptg t) -> t instanceof StringPtg;
    private final Predicate<Ptg> subtractPtg = (Ptg t) -> t instanceof SubtractPtg;
    private final Predicate<Ptg> tblPtg = (Ptg t) -> t instanceof TblPtg;
    private final Predicate<Ptg> unaryMinusPtg = (Ptg t) -> t instanceof UnaryMinusPtg;
    private final Predicate<Ptg> unaryPlusPtg = (Ptg t) -> t instanceof UnaryPlusPtg;
    private final Predicate<Ptg> unionPtg = (Ptg t) -> t instanceof UnionPtg;
    private final Predicate<Ptg> unknownPtg = (Ptg t) -> t instanceof UnknownPtg;
    private final Workbook workbook;

    public boolean verbose = false;
    public boolean errors = false;
    public boolean metadata = false;

    protected Class internalFormulaResultTypeClass;
    protected String formulaAddress;
    protected String formulaText;
    protected int formulaColumn;
    protected int formulaRow;
    protected String formulaRaw;
    protected int currentSheetIndex;
    protected String currentSheetName;
    protected String fileName;

    protected String creator;
    protected String description;
    protected String keywords;
    protected String title;
    protected String subject;
    protected String category;
    protected String author;
    protected String company;
    protected String template;
    protected String manager;
    protected String comment;
    private Sheet sheet;
    private EvaluationSheet evaluationSheet;
    private XSSFEvaluationWorkbook evaluationWorkbook;


    public AbstractParser(File file) throws InvalidFormatException, IOException {
        this(WorkbookFactory.create(file));
        this.fileName = file.getName();
    }

    public AbstractParser(Workbook workbook) {
        this.workbook = workbook;
        //metadata
        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) this.workbook;
        POIXMLProperties props = xssfWorkbook.getProperties();
        POIXMLProperties.CoreProperties coreProperties = props.getCoreProperties();

        this.creator = coreProperties.getCreator(); //get document creator
        this.description = coreProperties.getDescription(); //set Description
        this.keywords = coreProperties.getKeywords(); //set keywords
        this.title = coreProperties.getTitle(); //Title of the document
        this.subject = coreProperties.getSubject(); //Subject
        this.category = coreProperties.getCategory(); //cate


        POIXMLProperties.CustomProperties customProperties = props.getCustomProperties();
        customProperties.getProperty("Author");
        List<CTProperty> list = customProperties.getUnderlyingProperties().getPropertyList();
        for (CTProperty p : list) {

        }
        POIXMLProperties.ExtendedProperties extendedProperties = props.getExtendedProperties();
        this.company = extendedProperties.getUnderlyingProperties().getCompany();
        this.template = extendedProperties.getUnderlyingProperties().getTemplate();
        this.manager = extendedProperties.getUnderlyingProperties().getManager();

        this.evaluationWorkbook = XSSFEvaluationWorkbook.create((XSSFWorkbook) workbook);
        //System.out.println("Parse...");
    }

    public void verbose(String text) {
        if (this.verbose) System.out.println(text);
    }

    protected void err(String string, int row, int column) {

    }

    public void parse() {
        for (Sheet sht : this.workbook) {
            this.sheet = sht;
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

        comment = excelCell.getComment();

        formulaColumn = cell.getColumnIndex();
        formulaRow = cell.getRowIndex();
        internalFormulaResultTypeClass = excelCell.internalFormulaResultType();
        formulaAddress = HelperInternal.cellAddress(formulaRow, formulaColumn);
        formulaText = cell.getCellFormula();
        //formulaRaw = formulaAddress + " = " + formulaText;

        FormulaTokensInternal tokens = new FormulaTokensInternal(this.evaluationWorkbook, this.evaluationSheet);
        Ptg[] formulaPtgs = tokens.getFormulaTokens(formulaRow, formulaColumn);
        if (formulaPtgs == null) {
            System.err.println("ptgs empty or null for address " + formulaAddress);
            err("ptgs empty or null for address " + formulaAddress, formulaRow, formulaColumn);
            parseUDF(formulaText, formulaRow, formulaColumn);
            return;
        }
        Start start = parse(formulaPtgs, formulaRow, formulaColumn);
        start.setComment(comment);
        parseFormula(start);
    }

    protected abstract void parseUDF(String arguments, int formulaRow, int formulaColumn);


    private Start parse(Ptg[] ptgs, int row, int column) {
        Start start = null;
        parseFormulaInit();
        if (Ptg.doesFormulaReferToDeletedCell(ptgs)) doesFormulaReferToDeletedCell(row, column);
        for (Ptg ptg : ptgs) parse(ptg, row, column);
        start = parseFormulaPost(start, row, column);
        return start;
    }

    private void parse(Ptg p, int row, int column) {
        //verbose("parse: " + p.getClass().getSimpleName());

        Stream<WhatIf> stream = Stream.of(
                new WhatIf(p, arrayPtg, (Ptg t) -> parseArrayPtg((ArrayPtg) t)),
                new WhatIf(p, addPtg, (Ptg t) -> parseAdd()),
                new WhatIf(p, area3DPxg, (Ptg t) -> parseArea3DPxg((Area3DPxg) t)),
                new WhatIf(p, areaErrPtg, (Ptg t) -> parseAreaErrPtg((AreaErrPtg) t)),
                new WhatIf(p, areaNPtg, (Ptg t) -> parseAreaNPtg((AreaNPtg) t)),
                new WhatIf(p, areaPtg, (Ptg t) -> parseAreaPtg((AreaPtg) t)),
                new WhatIf(p, attrPtg, (Ptg t) -> parseAttrPtg((AttrPtg) t)),
                new WhatIf(p, boolPtg, t -> parseBool(((BoolPtg) t).getValue())),
                new WhatIf(p, concatPtg, t -> parseConcat()),
                new WhatIf(p, deleted3DPxg, (Ptg t) -> parseDeleted3DPxg((Deleted3DPxg) t)),
                new WhatIf(p, deletedArea3DPtg, (Ptg t) -> parseDeletedArea3DPtg((DeletedArea3DPtg) t)),
                new WhatIf(p, deletedRef3DPtg, (Ptg t) -> parseDeletedRef3DPtg((DeletedRef3DPtg) t)),
                new WhatIf(p, dividePtg, t -> parseDivide()),
                new WhatIf(p, equalPtg, t -> parseEqual()),
                new WhatIf(p, errPtg, (Ptg t) -> parseErrPtg((ErrPtg) t)),
                new WhatIf(p, expPtg, (Ptg t) -> parseExpPtg((ExpPtg) t)),
                new WhatIf(p, funcPtg, (Ptg t) -> parseFuncPtg((FuncPtg) t)),
                new WhatIf(p, funcVarPtg, (Ptg t) -> parseFuncVarPtg((FuncVarPtg) t)),
                new WhatIf(p, greaterEqualPtg, t -> parseGreaterEqual()),
                new WhatIf(p, greaterThanPtg, t -> parseGreaterThan()),
                new WhatIf(p, intersectionPtg, t -> parseIntersection()),
                new WhatIf(p, intPtg, t -> parseInt(((IntPtg) t).getValue())),
                new WhatIf(p, lessEqualPtg, t -> parseLessEqual()),
                new WhatIf(p, lessThanPtg, t -> parseLessThan()),
                new WhatIf(p, memAreaPtg, (Ptg t) -> parseMemAreaPtg((MemAreaPtg) t)),
                new WhatIf(p, memErrPtg, (Ptg t) -> parseMemErrPtg((MemErrPtg) t)),
                new WhatIf(p, memFuncPtg, (Ptg t) -> parseMemFuncPtg((MemFuncPtg) t)),
                new WhatIf(p, missingArgPtg, (Ptg t) -> parseMissingArgPtg((MissingArgPtg) t, row, column)),
                new WhatIf(p, multiplyPtg, t -> parseMultiply()),
                new WhatIf(p, namePtg, (Ptg t) -> parseNamePtg((NamePtg) t)),
                new WhatIf(p, nameXPxg, (Ptg t) -> parseNameXPxg((NameXPxg) t)),
                new WhatIf(p, notEqualPtg, t -> parseNotEqual()),
                new WhatIf(p, numberPtg, t -> parseNumber(((NumberPtg) t).getValue())),
                new WhatIf(p, parenthesisPtg, t -> parseParenthesis()),
                new WhatIf(p, percentPtg, t -> parsePercent()),
                new WhatIf(p, powerPtg, t -> parsePower()),
                new WhatIf(p, rangePtg, (Ptg t) -> parseRangePtg((RangePtg) t)),
                new WhatIf(p, ref3DPxg, (Ptg t) -> parseRef3DPxg((Ref3DPxg) t)),
                new WhatIf(p, refErrorPtg, (Ptg t) -> parseRefErrorPtg((RefErrorPtg) t)),
                new WhatIf(p, refNPtg, (Ptg t) -> parseRefNPtg((RefNPtg) t)),
                new WhatIf(p, refPtg, (Ptg t) -> parseRefPtg((RefPtg) t)),
                new WhatIf(p, stringPtg, (Ptg t) -> parseString(((StringPtg) t).getValue())),
                new WhatIf(p, subtractPtg, t -> parseSubtract()),
                new WhatIf(p, tblPtg, (Ptg t) -> parseTblPtg((TblPtg) t)),
                new WhatIf(p, unaryMinusPtg, (Ptg t) -> parseUnaryMinus()),
                new WhatIf(p, unaryPlusPtg, (Ptg t) -> parseUnaryPlus()),
                new WhatIf(p, unionPtg, t -> parseUnion()),
                new WhatIf(p, unknownPtg, (Ptg t) -> parseUnknownPtg((UnknownPtg) t))
        );

        stream.filter((WhatIf t) -> t.predicate.test(t.ptg)).forEach(t -> t.consumer.accept(t.ptg));
    }


    private void parseArrayPtg(ArrayPtg t) {
        Object[][] tokens = t.getTokenArrayValues();
        parseArray(tokens);
    }

    private void parseArea3DPtg(Area3DPtg t) {
        String area = t.format2DRefAsString();
        int externSheetIndex = t.getExternSheetIndex();
        parseArea3D(externSheetIndex, area);
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
        String area = t.format2DRefAsString();
        RangeInternal range = new RangeInternal(workbook, t.getSheetName(), t);
        parseArea3D(range.getFirstRow(), range.getFirstColumn(), range.getLastRow(), range.getLastColumn(), range.getValues(), sheetName, sheetIndex, area);
    }

    private void parseAttrPtg(AttrPtg t) {
        if (t.isSum()) sum();
    }

    private void parseAreaNPtg(AreaNPtg t) {
        RangeInternal range = new RangeInternal(workbook, sheet, t);
        parseAreaN(range.getValues(),
                range.getFirstRow(),
                range.getFirstColumn(),
                range.isFirstRowRelative(),
                range.isFirstColumnRelative(),
                range.getLastRow(),
                range.getLastColumn(),
                range.isLastRowRelative(),
                range.isLastColumnRelative());
    }

    private void parseAreaPtg(AreaPtg t) {
        RangeInternal range = new RangeInternal(workbook, sheet, t);
        parseArea(range.getValues(),
                range.getFirstRow(),
                range.getFirstColumn(),
                range.isFirstRowRelative(),
                range.isFirstColumnRelative(),
                range.getLastRow(),
                range.getLastColumn(),
                range.isLastRowRelative(),
                range.isLastColumnRelative());
    }

    private void parseErrPtg(ErrPtg t) {
        ErrInternal err = new ErrInternal(t);
        parseErr(err.text());
    }

    private void parseExpPtg(ExpPtg t) {
        parseExp();
    }

    private void parseFuncPtg(FuncPtg t) {
        String name = t.getName();
        int arity = t.getNumberOfOperands();
        parseFunc(name, arity);
    }

    private void parseFuncVarPtg(FuncVarPtg t) {
        String name = t.getName();
        int arity = t.getNumberOfOperands();
        parseFuncVar(name, arity);
    }

    private void parseMemAreaPtg(MemAreaPtg t) {
        System.err.println(t.toString());
        parseMemArea();
    }

    private void parseMemFuncPtg(MemFuncPtg t) {
        parseMemFunc();
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
        parseName(range.getFirstRow(), range.getFirstColumn(), range.getLastRow(), range.getLastColumn(), range.getValues(), name, range.getSheetName());
    }


    private void parseNameXPtg(NameXPtg nameXPtg) {
        String text = nameXPtg.toString();
        parseNameX(text);
    }

    private void parseNameXPxg(NameXPxg t) {
        String text = t.toString();
        parseNameX(text);
    }

    private void parseRangePtg(RangePtg t) {
        parseRange(t.toString());
    }

    private void parseRef3DPxg(Ref3DPxg t) {
        int extWorkbookNumber = t.getExternalWorkbookNumber();
        String sheet_ = t.getSheetName();
        String area = t.format2DRefAsString();
        parseRef3D(extWorkbookNumber, sheet_, area);
    }

    /*private void parseRef3DPtg(Ref3DPtg ref3DPtg) {
        int ext = ref3DPtg.getExternSheetIndex();
        String sheetName = this.initializeEvaluationWorkbook.getSheetName(ext);
        String area = ref3DPtg.format2DRefAsString();
        parseRef3DPtg(sheetName, area);
    }*/
    private void parseRefNPtg(RefNPtg t) {

        parseRefN(t.toString());
    }

    private void parseRefPtg(RefPtg t) {
        int ri = t.getRow();
        int ci = t.getColumn();
        boolean rowRelative = t.isRowRelative();
        boolean colRelative = t.isColRelative();
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
        parseRef(ri, ci, rowRelative, colRelative, rowNotNull, value, comment);
    }

    private void parseTblPtg(TblPtg t) {
        parseTbl(t.toString());
    }

    //<editor-fold defaultstate="collapsed" desc="Missing And Error">
    private void parseRefErrorPtg(RefErrorPtg t) {
        String text = t.toString();
        parseRefError(text);
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

    protected void parseMissingArgPtg(MissingArgPtg t, int row, int column) {
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
//</editor-fold>      

    //<editor-fold defaultstate="collapsed" desc="Abstract-Methods">
    protected abstract void parseFormula(Start start);

    protected abstract void parseMissingArguments(int row, int column);

    protected abstract void parseArray(Object[][] array);

    protected abstract void parseAdd();

    protected abstract void parseArea3D(int externSheetIndex, String area);

    protected abstract void parseArea3D(int FirstRow, int FirstColumn, int LastRow, int LastColumn, List<Object> list, String sheetName, int sheetIndex, String area);

    protected abstract void parseAreaN(List<Object> list, int firstRow, int firstColumn, boolean isFirstRowRelative, boolean isFirstColRelative, int lastRow, int lastColumn, boolean isLastRowRelative, boolean isLastColRelative);

    protected abstract void sum();

    protected abstract void parseArea(List<Object> list, int firstRow, int firstColumn, boolean isFirstRowRelative, boolean isFirstColRelative, int lastRow, int lastColumn, boolean isLastRowRelative, boolean isLastColRelative);

    protected abstract void parseBool(Boolean bool);

    protected abstract void parseConcat();

    protected abstract void parseDivide();

    protected abstract void parseEqual();

    protected abstract void parseErr(String text);

    protected abstract void parseExp();

    protected abstract void parseFunc(String name, int arity);

    protected abstract void parseFuncVar(String name, int arity);

    protected abstract void parseGreaterEqual();

    protected abstract void parseGreaterThan();

    protected abstract void parseIntersection();

    protected abstract void parseInt(Integer value);

    protected abstract void parseLessEqual();

    protected abstract void parseLessThan();

    protected abstract void parseMemArea();

    protected abstract void parseMemFunc();

    protected abstract void parseMultiply();

    protected abstract void parseName(int firstRow, int firstColumn, int lastRow, int lastColumn, List<Object> cells, String name, String sheetName);

    protected abstract void parseNameX(String name);

    protected abstract void parseNotEqual();

    protected abstract void parseNumber(Double value);

    protected abstract void parseParenthesis();

    protected abstract void parsePercent();

    protected abstract void parsePower();

    protected abstract void parseRange(String text);

    protected abstract void parseRef3D(int ext, String sheet, String area);

    //protected abstract void parseRef3DPtg(String sheetName, String area);
    protected abstract void parseRefError(String text);

    protected abstract void parseRefN(String text);

    protected abstract void parseRef(int ri, int ci, boolean rowRelative, boolean colRelative, boolean rowNotNull, Object value, String comment);

    protected abstract void parseString(String string);

    protected abstract void parseSubtract();

    protected abstract void parseTbl(String text);

    protected abstract void parseUnaryMinus();

    protected abstract void parseUnaryPlus();

    protected abstract void parseUnion();

    protected abstract void doesFormulaReferToDeletedCell(int row, int column);

    protected abstract void parseFormulaInit();

    protected abstract Start parseFormulaPost(Start start, int row, int column);
//</editor-fold>  

    public String getFileName() {
        return fileName;
    }

    // 3DPxg is XSSF
    // 3DPtg is HSSF
    class WhatIf {

        final Ptg ptg;
        final Predicate<Ptg> predicate;
        final Consumer<Ptg> consumer;

        public WhatIf(Ptg ptg, Predicate<Ptg> predicate, Consumer<Ptg> consumer) {
            this.ptg = ptg;
            this.predicate = predicate;
            this.consumer = consumer;
        }
    }

    //BEGIN
    //END


}
