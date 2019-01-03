;
; Efesto - Excel Formula Extractor System and Topological Ordering algorithm.
; Copyright (C) 2017 Massimo Caliman mcaliman@caliman.biz
;
; This program is free software: you can redistribute it and/or modify
; it under the terms of the GNU Affero General Public License as published
; by the Free Software Foundation, either version 3 of the License, or
; (at your option) any later version.
;
; This program is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
; GNU Affero General Public License for more details.
;
; You should have received a copy of the GNU Affero General Public License
; along with this program.  If not, see <https://www.gnu.org/licenses/>.
;
; If AGPL Version 3.0 terms are incompatible with your use of
; Efesto, alternative license terms are available from Massimo Caliman
; please direct inquiries about Efesto licensing to mcaliman@caliman.biz
;

;
; Text File: test/03-range-index.clj
; Excel File: 03-range-index.xlsx
; Excel Formulas Number: 1
; Elapsed Time (parsing + topological sort): 2 s. or 0 min.
; creator:null
; description:null
; keywords:null
; title:null
; subject:null
; category:null
(def A1:B6 A1:B6)
(def A10 RANGE!A10 = INDEX (A1:B6, 2, 2))
