(ns horizontal-maker.odt-test
  (:require [clojure.java.io :as io]
            [clojure.test :refer :all]
            [horizontal-maker.formats :as formats]
            [horizontal-maker.test-helpers :as h]
            [horizontal-maker.vertical :as vertical])
  (:import [org.odftoolkit.odfdom.doc OdfTextDocument]
           [org.odftoolkit.odfdom.doc.table OdfTable]))

(deftest test-make-vertical
  (let [path (h/temp-file "vertical.odt")
        parsed (vertical/generate-vertical-parsed h/sample-rows "Romans")]
    (formats/make-vertical! path h/sample-rows "Romans")
    (let [doc (OdfTextDocument/loadDocument (io/file path))
          tables (vec (.getTableList doc))]
      (try
        (is (= (count parsed) (count tables)))
        (doseq [[[name paragraphs ref] ^OdfTable table] (map vector parsed tables)]
          (is (= (+ 3 (count paragraphs)) (.getRowCount table)))
          (is (= 3 (.getColumnCount table)))
          (is (= name (.getStringValue (.getCellByPosition table 1 0))))
          (is (= ref (.getStringValue (.getCellByPosition table 2 0))))
          (is (= paragraphs (map #(.getStringValue (.getCellByPosition table 1 (int (inc %))))
                                 (range (count paragraphs)))))
          (is (= "Summary:" (.getStringValue (.getCellByPosition table 1 (inc (count paragraphs))))))
          (is (= "Application:" (.getStringValue (.getCellByPosition table 0 (+ 2 (count paragraphs))))))
          (is (= (inc (count paragraphs))
                 (.getTableNumberRowsSpannedAttribute (.getOdfElement (.getCellByPosition table 0 1))))))
        (is (= "Segment 1" (.getStringValue (.getCellByPosition (first tables) 0 0))))
        (finally (.close doc))))))
