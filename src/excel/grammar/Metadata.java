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

package excel.grammar;

/**
 * @author Massimo Caliman
 */
public final class Metadata extends Start {
    private String name;
    private Object value;

    public Metadata(String name, Object value) {
        this.name = name;
        this.value = value;
    }

    public String getName() {
        return name;
    }

    public Object getValue() {
        return value;
    }

    @Override
    public String toString() {
        if (emptyName(name) || emptyValue(value)) return "''";
        if (emptyValue(value)) return "";
        return "'' " + name + " : " + value.toString();
    }

    public String toString(boolean address) {
        if (emptyName(name) || emptyValue(value)) return "''";
        if (emptyValue(value)) return "";
        return "'' " + name + " : " + value.toString();
    }

    private boolean emptyName(String variable) {
        return variable == null || variable.trim().length() == 0;
    }

    private boolean emptyValue(Object value) {
        return value == null || (value instanceof String && ((String) value).trim().length() == 0);
    }
}
