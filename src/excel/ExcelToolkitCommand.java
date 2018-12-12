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
        parser.parse();
        elapsed = System.currentTimeMillis() - t;
    }

    public void writer(@NotNull String filename) throws IOException {
        StartList list = parser.getList();
        try (Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(filename), StandardCharsets.UTF_8))) {
            writer.write("'' Text File: " + filename + '\n');
            writer.write("'' Excel File: " + parser.getFileName()  + '\n');
            writer.write("'' Elapsed Time (parsing + topological sort): " + (elapsed / 1000 + " s. or " + (elapsed / 1000 / 60) + " min.") + '\n');
            //writer.write("'' creator:" + parser.getCreator()+'\n');
            //writer.write("'' description:"+ parser.getDescription()+'\n');
            //writer.write("'' keywords:"+parser.getKeywords()+'\n');
            //writer.write("'' title:"+parser.getTitle()+'\n');
            //writer.write("'' subject:"+parser.getSubject()+'\n');
            //writer.write("'' category:"+parser.getCategory()+'\n');
            //writer.write("'' company:"+parser.getCompany()+'\n');
            //writer.write("'' template:"+parser.getTemplate()+'\n');
            //writer.write("'' manager:"+parser.getManager()+'\n');
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
            //System.out.println("" + start.getClass().getSimpleName() + " : " + start.getAddress(true) + " = " + start.toString(false));
    }

    private StartList getStartList() {
        return parser.getList();
    }

    public boolean test(int offset, String... text) {
        return getStartList().test(offset, text);
    }

    public void toFunctional(){
        for (Start start : getStartList()){
            //System.out.println("" + start.getClass().getSimpleName() + " : " );
            if(start instanceof ToFunctional) {
                System.out.println( start.getSheetName() + start.getAddress(false) + " = " +  ((ToFunctional) start).toFuctional() );
            }else {
                System.out.println(start.toString(true));
            }
        }
    }

}
