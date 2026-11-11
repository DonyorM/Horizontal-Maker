(ns horizontal-maker.ods-test
  (:require [clojure.java.io :as io]
            [clojure.test :refer :all]
            [horizontal-maker.core :as core]
            [horizontal-maker.formats :as formats]
            [horizontal-maker.sheet :as sheet]
            [horizontal-maker.test-helpers :as h])
  (:import [org.odftoolkit.odfdom.doc OdfSpreadsheetDocument]
           [org.odftoolkit.odfdom.doc.table OdfTable]))

(defn- write-sample!
  "Writes the sample sheet as an .ods, leaving blank cells untouched."
  [path]
  (let [doc (OdfSpreadsheetDocument/newSpreadsheetDocument)
        table (first (.getTableList doc))]
    (.setTableName table "Input")
    (doseq [[r row] (map-indexed vector h/sample-sheet)
            [c v] (map-indexed vector row)
            :when (some? v)]
      (let [cell (.getCellByPosition table (int c) (int r))]
        (if (number? v)
          (.setDoubleValue cell (double v))
          (.setStringValue cell v))))
    (.save doc (io/file path))
    (.close doc)
    path))

(defn- table-names [path]
  (let [doc (OdfSpreadsheetDocument/loadDocument (io/file path))]
    (try
      (mapv #(.getTableName ^OdfTable %) (.getTableList doc))
      (finally (.close doc)))))

(deftest test-read-rows
  (is (= h/sample-rows (sheet/read-rows (write-sample! (h/temp-file "sample.ods"))))))

(deftest test-make-hzd
  (let [path (write-sample! (h/temp-file "sample.ods"))
        layout (core/hzd-layout h/sample-rows)]
    (formats/make-hzd! path)
    ;; Generating twice must replace, not duplicate, the HZD sheet
    (formats/make-hzd! path)
    (is (= ["Input" "HZD"] (table-names path)))
    (let [doc (OdfSpreadsheetDocument/loadDocument (io/file path))
          hzd (.getTableByName doc "HZD")
          cell (fn [col row] (.getCellByPosition hzd (int col) (int row)))]
      (try
        (is (= (count (:rows layout)) (.getRowCount hzd)))
        (is (= "Sect." (.getStringValue (cell 0 5))))
        (is (= "Section 1" (.getStringValue (cell 1 5))))
        (is (= "1:1" (.getStringValue (cell 1 9))))
        (is (= 15.0 (.getDoubleValue (cell 3 7))))
        (is (= 5 (.getTableNumberColumnsSpannedAttribute (.getOdfElement (cell 1 5)))))
        (is (pos? (.getWidth (.getColumnByIndex hzd 1))))
        (is (pos? (.getHeight (.getRowByIndex hzd 9))))
        (is (not= (.getStyleName (cell 1 9)) (.getStyleName (cell 1 5))))
        (finally (.close doc))))
    (testing "the input sheet is untouched"
      (is (= h/sample-rows (sheet/read-rows path))))))
