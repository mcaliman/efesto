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

package excel.grammar.formula.reference;

/**
 * SHEET The name of a worksheet
 * Priority: 5
 * Sheet characters
 * Any character except ' * [ ] \ : / ? ( ) ; { } # " = < > & + - * / ^ % , ␣
 *
 * @author Massimo Caliman
 */
public final class SHEET extends Prefix {

    private final String name;

    public SHEET(String name) {
        this.name = name;
    }

    public String toString() {
        return name;
    }
}
