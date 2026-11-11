(ns horizontal-maker.sheet-test
  (:require [clojure.test :refer :all]
            [horizontal-maker.sheet :as sheet]))

(deftest test-extension
  (is (= :xlsx (sheet/extension "/tmp/Book.xlsx")))
  (is (= :ods (sheet/extension "Book.ODS")))
  (is (= :odt (sheet/extension (java.io.File. "dir.v2" "vertical.odt"))))
  (is (nil? (sheet/extension "noext")))
  (is (nil? (sheet/extension ".hidden"))))

(deftest test-normalize-rows
  (is (= [[nil 1.0 "Four words" nil "Segment 1" "Section 1" nil]
          ["2" 5.0 "Para 2" 15.0 nil nil nil]]
         (sheet/normalize-rows [["Chapter" "Verse" "Paragraph Title" "End" "Segment" "Section" "Division"]
                                [nil 1 "Four words" "" "Segment 1" "Section 1"]
                                [2.0 5.0 "Para 2" 15 "" "" "" "extra"]
                                [nil nil nil nil nil nil nil]
                                ["" "x" "not a verse"]])))
  (is (= [] (sheet/normalize-rows []))))

(deftest test-unsupported-format
  (is (thrown-with-msg? clojure.lang.ExceptionInfo #"Unsupported spreadsheet format"
                        (sheet/read-rows "input.csv")))
  (is (thrown-with-msg? clojure.lang.ExceptionInfo #"Unsupported spreadsheet format"
                        (sheet/write-hzd! "input.csv" {}))))
