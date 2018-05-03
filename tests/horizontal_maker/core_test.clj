(ns horizontal-maker.core-test
  (:require [clojure.test :refer :all]
            [horizontal-maker.core :refer :all])
  (:import org.apache.poi.ss.util.CellRangeAddress))

(def example-input-1 [[1 "Paragraph 1" "" "Appoint"    "Section 1" "Division 1"]
                      [5 "Para 2"      "" ""           ""           ""]
                      [16 "Para 3"     20 "Elders"     ""           ""]
                      [1 "Para 4"      ""  ""          ""           ""]
                      [15 "Para 5"     ""  "More"      "Section 2"  ""]
                      [21 "Para 6"     30  ""          ""           ""]
                      [1 "Para 7"      ""  "Great"     ""           ""]
                      [10 "Para 8"     ""  ""          ""           ""]
                      [18 "Para 9"     ""  "Fantas"    "Section 3"  ""]
                      [21 "Para 10"     30  ""          ""           ""]
                      [1 "Para 11"      ""  "Cool"     "Section 4"   "Division 2"]
                      [10 "Para 12"     ""  ""          ""           ""]
                      [18 "Para 13"     25  "Amaze"    "Section 5"  ""]
                      [2 "Para 14"     ""  ""          ""           ""]
                      [18 "Para 15"     22 ""          ""           ""]])

(def example-input-no-div [[1 "Paragraph 1" "" "Appoint"    "Section 1" ""]
                           [5 "Para 2"      "" ""           ""           ""]
                           [16 "Para 3"     20 "Elders"     ""           ""]
                           [1 "Para 4"      ""  ""          ""           ""]
                           [15 "Para 5"     ""  "More"      "Section 2"  ""]
                           [21 "Para 6"     30  ""          ""           ""]
                           [1 "Para 7"      ""  "Great"     ""           ""]
                           [10 "Para 8"     15          ""           ""]])

(def example-parsed-no-div [["Book:"] ["Title:"] ["Key Verse:"]
                            [""][""]
                            ["Sect." "Section 1" "" "" "" "Section 2" "" "" ""]
                            ["Segm." "Appoint" "" "Elders" "" "More" "" "Great" ""]
                            ["" "" "" 20 "" "" 30 "" 15]
                            ["Paragr." "Paragraph 1" "Para 2" "Para 3" "Para 4" "Para 5" "Para 6" "Para 7" "Para 8"]
                            ["" "1:1" "1:5" "1:16" "2:1" "2:15" "2:21" "3:1" "3:10"]])

(def example-parsed-1 [["Book:"] ["Title:"] ["Key Verse:"]
                    [""] [""]
                    ["Divis." "Division 1" "" "" "" "" "" "" "" "" "" "Division 2" "" "" "" ""]
                    ["Sect." "Section 1" "" "" "" "Section 2" "" "" "" "Section 3" "" "Section 4" "" "Section 5" "" ""]
                    ["Segm." "Appoint" "" "Elders" "" "More" "" "Great" "" "Fantas" "" "Cool" "" "Amaze" "" ""]
                    ["" "" "" 20 "" "" 30 "" "" "" 30 "" "" 25 "" 22]
                    ["" "Paragraph 1" "Para 2" "Para 3" "Para 4" "Para 5" "Para 6" "Para 7" "Para 8" "Para 9" "Para 10" "Para 11" "Para 12" "Para 13" "Para 14" "Para 15"]
                    ["" 1 5 16 1 15 21 1 10 18 21 1 10 18 2 18]])

(deftest test-get-end-references
  (is (= ["" 20 30 "" 30 "" 22]
         (get-end-references example-input-1 2 3))))

(deftest test-get-start-references
  (is (= ["1:1" "1:16" "2:15" "3:1" "3:18" "4:1" "4:18"]
         (get-start-references example-input-1 0 3))))

(deftest test-cell-range-address
  (is (.equals (cell-range-address 0 1 1 3)
               (CellRangeAddress. 0 1 1 3))))

(deftest test-get-column
  (is (= ["Division 1" "" "" "" "" "" "" "" "" "" "Division 2" "" "" "" ""]
         (get-column example-input-1 5))))

(deftest test-get-merge-range
  (is (= [[1 10] [11 15]]
         (get-merge-range (nth example-parsed-1 5)))))

(deftest test-create-spreadsheet-vector
  (is (= [["Book:"]
          ["Title:"]
          ["Key Verse:"]
          [""]
          [""]
          ["Divis.""Division 1" "" "" "" "" "Division 2"  ""]
          ["Sect." "Section 1" "" "Section 2" "" "Section 3" "Section 4" "Section 5"]
          ["" "" 20 30 "" 30 "" 22]
          ["Segm." "Appoint" "Elders" "More" "Great" "Fantas" "Cool" "Amaze"]
          ["" "1:1" "1:16" "2:15" "3:1" "3:18" "4:1" "4:18"]]
         (create-spreadsheet-vector example-input-1)))
  (is (= example-parsed-no-div
         (create-spreadsheet-vector example-input-no-div))))

(deftest test-calcualte-width
  (is (< (Math/abs (apply - TOTAL-WIDTH (calculate-width example-input-no-div)))
         50))
  (is (= (count example-input-no-div)
         (count (calculate-width example-input-no-div))))
  (is (< (Math/abs (apply - TOTAL-WIDTH (calculate-width example-input-1)))
         50))
  (is (= (count (remove empty? (map #(nth % SEG-IDX) example-input-1)))
         (count (calculate-width example-input-1)))))
