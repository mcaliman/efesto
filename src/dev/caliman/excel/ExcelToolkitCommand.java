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

package dev.caliman.excel;

import dev.caliman.excel.grammar.Comment;
import dev.caliman.excel.grammar.Start;
import dev.caliman.excel.parser.Parser;
import dev.caliman.excel.parser.StartList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jetbrains.annotations.NotNull;

import java.io.*;
import java.nio.charset.StandardCharsets;

public class ExcelToolkitCommand implements ToolkitCommand {

    private Parser parser;

    private long elapsed = 0;

    public ExcelToolkitCommand(@NotNull String name) throws IOException, InvalidFormatException {
        ToolkitOptions options = new ToolkitOptions();
        parser = new Parser(name);
        parser.verbose = options.isVerbose();
        parser.metadata = options.isMetadata();
    }

    public ExcelToolkitCommand(@NotNull String name, ToolkitOptions options) throws IOException, InvalidFormatException {
        parser = new Parser(name);
        parser.verbose = options.isVerbose();
        parser.metadata = options.isMetadata();
    }


    public void execute() {
        long t = System.currentTimeMillis();
        this.parser.parse();
        this.elapsed = System.currentTimeMillis() - t;
    }

    public void writerFormula(@NotNull String filename) throws IOException {
        StartList list = parser.getList();
        try (Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(filename), StandardCharsets.UTF_8))) {
            writer.write("'' \n");
            writer.write("'' Text File: " + filename + '\n');
            writer.write("'' Excel File: " + parser.getFileName() + '\n');
            writer.write("'' Excel Formulas Number: " + parser.getCounterFormulas() + '\n');
            writer.write("'' Elapsed Time (parsing + topological sort): " + (elapsed / 1000 + " s. or " + (elapsed / 1000 / 60) + " min.") + '\n');
            writer.write("'' creator:" + parser.getCreator() + '\n');
            writer.write("'' description:" + parser.getDescription() + '\n');
            writer.write("'' keywords:" + parser.getKeywords() + '\n');
            writer.write("'' title:" + parser.getTitle() + '\n');
            writer.write("'' subject:" + parser.getSubject() + '\n');
            writer.write("'' category:" + parser.getCategory() + '\n');
            for (Start start : list) {
                Comment comment = start.getComment();
                //if (comment != null && comment.trim().length() > 0)
                if (comment != null) writer.write(comment.toString());
                try {
                    //writer.write(start.id() + " = " + start.toFormula());
                    writer.write(start.id() + " = " + start.toString());
                } catch (Exception e) {
                    writer.write("'' Erron when compile " + start.id());
                }
                writer.write("\n");
            }
        }
    }

    @Deprecated
    public void writerLanguage(@NotNull String filename) throws IOException {
        StartList list = parser.getList();
        try (Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(filename), StandardCharsets.UTF_8))) {
            writer.write("; \n");
            writer.write("; Text File: " + filename + '\n');
            writer.write("; Excel File: " + parser.getFileName() + '\n');
            writer.write("; Excel Formulas Number: " + parser.getCounterFormulas() + '\n');
            writer.write("; Elapsed Time (parsing + topological sort): " + (elapsed / 1000 + " s. or " + (elapsed / 1000 / 60) + " min.") + '\n');
            writer.write("; creator:" + parser.getCreator() + '\n');
            writer.write("; description:" + parser.getDescription() + '\n');
            writer.write("; keywords:" + parser.getKeywords() + '\n');
            writer.write("; title:" + parser.getTitle() + '\n');
            writer.write("; subject:" + parser.getSubject() + '\n');
            writer.write("; category:" + parser.getCategory() + '\n');
            for (Start start : list) {
                Comment comment = start.getComment();
                if (comment != null) writer.write("; " + comment.toString());
                writer.write("(def " + start.id() + " " + start.toLanguage() + ")");
                writer.write("\n");
            }
        }
    }



    private StartList getStartList() {
        return parser.getList();
    }

    public boolean testToFormula(int offset, String... text) {
        return getStartList().testToFunctional(offset, text);
    }

    public void toFormula() {
        for (Start start : getStartList()) {
            try {
                if (start != null)
                    //System.out.println(start.id() + " = " + start.toFormula());
                    System.out.println(start.id() + " = " + start.toString());
            } catch (Exception e) {
                System.err.println("Error when transpile " + start.id());
            }
        }
    }

    public void toLanguage() {
        for (Start start : getStartList()) {
            if (start != null) {
                System.out.print("<" + start.getClass().getSimpleName() + ">");
                if (start instanceof ToLanguage)
                    System.out.println("[ToLanguage]:: (def " + start.id() + " " + start.toLanguage() + ")");
                else
                    System.out.println(start.id() + " = " + start.toFormula());
            }
        }
    }

}
