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

package efesto.internals;

import org.apache.poi.ss.formula.ptg.ErrPtg;

import static org.apache.poi.ss.formula.ptg.ErrPtg.*;

public final class ErrInternal {

    public final static String ERROR_NULL_INTERSECTION = "#NULL!";
    public final static String ERROR_DIV_ZERO = "#DIV/0!";
    public final static String ERROR_VALUE_INVALID = "#VALUE!";
    public final static String ERROR_REF_INVALID = "#REF!";
    public final static String ERROR_NAME_INVALID = "#NAME?";
    public final static String ERROR_NUM_ERROR = "#NUM!";
    public final static String ERROR_N_A = "#N/A";

    private final ErrPtg t;

    public ErrInternal(ErrPtg t) {
        this.t = t;
    }

    public String text() {
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
