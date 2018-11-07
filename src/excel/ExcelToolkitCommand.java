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

package excel;

import excel.grammar.Start;
import excel.parser.Parser;
import excel.parser.StartList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.*;
import java.nio.charset.StandardCharsets;

public class ExcelToolkitCommand implements ToolkitCommand {

    private Parser parser;

    public ExcelToolkitCommand(String name) throws IOException, InvalidFormatException {
        ToolkitOptions options = new ToolkitOptions();
        parser = new Parser(name);
        parser.verbose = options.isVerbose();
        parser.metadata = options.isMetadata();
    }

    public ExcelToolkitCommand(String name, ToolkitOptions options) throws IOException, InvalidFormatException {
        parser = new Parser(name);
        parser.verbose = options.isVerbose();
        parser.metadata = options.isMetadata();
    }


    @Override
    public void execute() {
        parser.parse();
    }

    public void writer(String filename) throws IOException {
        StartList list = parser.getList();
        try (Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(filename), StandardCharsets.UTF_8))) {
            for (Start start : list) {
                String comment = start.getComment();
                if (comment != null && comment.trim().length() > 0) writer.write("'' " + comment + "\n");
                writer.write(start.toString(true));
                writer.write("\n");
            }
        }
    }


    public void print() {
        for (Start start : getStartList())
            System.out.println("" + start.getClass().getSimpleName() + " : " + start.toString(true));
    }

    public void transpiler() {

    }

    public StartList getStartList() {
        return parser.getList();
    }

    public boolean test(int offset, String... text) {
        return getStartList().test(offset, text);
    }

}
