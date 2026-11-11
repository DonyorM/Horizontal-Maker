(ns horizontal-maker.excel-test
  (:require [clojure.test :refer :all]
            [dk.ative.docjure.spreadsheet :as sp]
            [horizontal-maker.core :as core]
            [horizontal-maker.excel :as excel]
            [horizontal-maker.formats :as formats]
            [horizontal-maker.sheet :as sheet]
            [horizontal-maker.test-helpers :as h])
  (:import [org.apache.poi.ss.util CellRangeAddress]))

(defn- write-sample!
  "Writes the sample sheet, leaving blank cells uncreated."
  [path]
  (let [wb (sp/create-workbook "Input" [])
        sheet (sp/select-sheet "Input" wb)]
    (doseq [[r row] (map-indexed vector h/sample-sheet)
            :let [poi-row (.createRow sheet (int r))]
            [c v] (map-indexed vector row)
            :when (some? v)]
      (sp/set-cell! (.createCell poi-row (int c)) v))
    (sp/save-workbook! path wb)
    path))

(deftest test-cell-range-address
  (is (.equals (excel/cell-range-address 0 1 1 3)
               (CellRangeAddress. 0 1 1 3))))

(deftest test-read-rows
  (is (= h/sample-rows (sheet/read-rows (write-sample! (h/temp-file "sample.xlsx")))))
  (is (= h/sample-rows (sheet/read-rows (write-sample! (h/temp-file "sample.xls"))))))

(deftest test-make-hzd
  (let [path (write-sample! (h/temp-file "sample.xlsx"))
        layout (core/hzd-layout h/sample-rows)]
    (formats/make-hzd! path)
    ;; Generating twice must replace, not duplicate, the HZD sheet
    (formats/make-hzd! path)
    (let [wb (sp/load-workbook path)
          hzd (sp/select-sheet "HZD" wb)]
      (is (= ["Input" "HZD"] (map sp/sheet-name (sp/sheet-seq wb))))
      (is (= (count (:rows layout)) (count (sp/row-seq hzd))))
      (is (= "Sect." (sp/read-cell (sp/select-cell "A6" hzd))))
      (is (= "Section 1" (sp/read-cell (sp/select-cell "B6" hzd))))
      (is (= "1:1" (sp/read-cell (sp/select-cell "B10" hzd))))
      (is (= (count (:merges layout)) (.getNumMergedRegions hzd)))
      (is (= (first (:col-widths layout)) (.getColumnWidth hzd 1)))
      (is (= 90 (.getRotation (.getCellStyle (sp/select-cell "B9" hzd)))))
      (is (true? (.getBold (.getFontAt wb (.getFontIndex (.getCellStyle (sp/select-cell "A1" hzd))))))))))
