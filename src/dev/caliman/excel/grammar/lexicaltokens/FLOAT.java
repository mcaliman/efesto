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

package dev.caliman.excel.grammar.lexicaltokens;

import dev.caliman.excel.grammar.annotations.LexicalTokens;

import java.util.Objects;

/**
 * @author Massimo Caliman
 */
@LexicalTokens(name = "FLOAT",
        description = "An integer, floating point or scientific notation number literal",
        content = "[0-9]+ ,? [0-9]* (e [0-9]+)?", priority = 0)
public final class FLOAT extends dev.caliman.excel.grammar.formula.constant.Number {

    private final Double value;

    public FLOAT(Double value) {
        this.value = value;
    }


    public boolean isTerminal() {
        return true;
    }

    @Override
    public int hashCode() {
        int hash = 7;
        hash = 23 * hash + Objects.hashCode(this.value);
        return hash;
    }

    @Override
    public boolean equals(Object obj) {
        if(this == obj)
            return true;
        if(obj == null)
            return false;
        if(getClass() != obj.getClass())
            return false;
        final FLOAT other = (FLOAT) obj;
        return Objects.equals(this.value, other.value);
    }

    @Override
    public String toString() {
        return value.toString();
    }

}
