(ns horizontal-maker.core-test
  (:require [clojure.test :refer :all]
            [horizontal-maker.constants :as con]
            [horizontal-maker.core :refer :all]))

;; Input rows follow the sheet layout:
;; [chapter verse title end segment section division]

(def example-input-1 [["1" 1 "Paragraph 1" "" "Appoint"    "Section 1" "Division 1"]
                      [""  5 "Para 2"      "" ""           ""           ""]
                      [""  16 "Para 3"     20 "Elders"     ""           ""]
                      ["2" 1 "Para 4"      ""  ""          ""           ""]
                      [""  15 "Para 5"     ""  "More"      "Section 2"  ""]
                      [""  21 "Para 6"     30  ""          ""           ""]
                      ["3" 1 "Para 7"      ""  "Great"     ""           ""]
                      [""  10 "Para 8"     ""  ""          ""           ""]
                      [""  18 "Para 9"     ""  "Fantas"    "Section 3"  ""]
                      [""  21 "Para 10"    30  ""          ""           ""]
                      ["4" 1 "Para 11"     ""  "Cool"      "Section 4"  "Division 2"]
                      [""  10 "Para 12"    ""  ""          ""           ""]
                      [""  18 "Para 13"    25  "Amaze"     "Section 5"  ""]
                      ["5" 2 "Para 14"     ""  ""          ""           ""]
                      [""  18 "Para 15"    22 ""           ""           ""]])

(def example-input-no-div [["1" 1 "Paragraph 1" "" "Appoint"    "Section 1" ""]
                           [""  5 "Para 2"      "" ""           ""           ""]
                           [""  16 "Para 3"     20 "Elders"     ""           ""]
                           ["2" 1 "Para 4"      ""  ""          ""           ""]
                           [""  15 "Para 5"     ""  "More"      "Section 2"  ""]
                           [""  21 "Para 6"     30  ""          ""           ""]
                           ["3" 1 "Para 7"      ""  "Great"     ""           ""]
                           [""  10 "Para 8"     15  ""          ""           ""]])

(def example-input-no-seg [["1" 1 "Paragraph 1" "" "" "" ""]
                           [""  5 "Para 2"      "" "" "" ""]
                           [""  16 "Para 3"     20 "" "" ""]
                           ["2" 1 "Para 4"      15 "" "" ""]])

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
         (get-end-references example-input-1 con/END-IDX con/SEG-IDX))))

(deftest test-get-start-references
  (is (= ["1:1" "1:16" "2:15" "3:1" "3:18" "4:1" "4:18"]
         (get-start-references example-input-1 con/START-IDX con/SEG-IDX))))

(deftest test-get-column
  (is (= ["Division 1" "" "" "" "" "" "" "" "" "" "Division 2" "" "" "" ""]
         (get-column example-input-1 con/DIV-IDX))))

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
         (create-spreadsheet-vector example-input-no-div)))
  (is (= [["Book:"]
          ["Title:"]
          ["Key Verse:"]
          [""]
          [""]
          ["" "" "" 20 15]
          ["Paragr." "Paragraph 1" "Para 2" "Para 3" "Para 4"]
          ["" "1:1" "1:5" "1:16" "2:1"]]
         (create-spreadsheet-vector example-input-no-seg))))

(deftest test-blank-cells-as-nil
  (testing "readers may hand over nil for blank cells"
    (let [nil-input (mapv (fn [row] (mapv #(if (= "" %) nil %) row)) example-input-no-div)]
      (is (= (map #(map (fn [v] (if (nil? v) "" v)) %) (create-spreadsheet-vector nil-input))
             (map #(map (fn [v] (if (nil? v) "" v)) %) (create-spreadsheet-vector example-input-no-div))))
      (is (= (calculate-width example-input-no-div)
             (calculate-width nil-input))))))

(deftest test-calculate-width
  (is (< (Math/abs (apply - con/TOTAL-WIDTH (calculate-width example-input-no-div)))
         50))
  (is (= (count example-input-no-div)
         (count (calculate-width example-input-no-div))))
  (is (< (Math/abs (apply - con/TOTAL-WIDTH (calculate-width example-input-1)))
         50))
  (is (= (count (remove empty? (map #(nth % con/SEG-IDX) example-input-1)))
         (count (calculate-width example-input-1))))
  (is (< (Math/abs (apply - 10000 (calculate-width example-input-1 10000)))
         50)))

(deftest test-merge-ranges
  (is (= [[5 1 10] [5 11 15]
          [6 1 4] [6 5 8] [6 9 10] [6 11 12] [6 13 15]
          [7 1 2] [7 3 4] [7 5 6] [7 7 8] [7 9 10] [7 11 12] [7 13 15]]
         (merge-ranges example-parsed-1)))
  (is (= [[5 1 4] [5 5 8]
          [6 1 2] [6 3 4] [6 5 6] [6 7 8]]
         (merge-ranges example-parsed-no-div))))

(deftest test-row-heights
  (is (= {8 con/PARA-ROW-HEIGHT-PT 9 con/REF-ROW-HEIGHT-PT}
         (row-heights example-parsed-no-div)))
  (is (= {9 con/PARA-ROW-HEIGHT-PT 10 con/REF-ROW-HEIGHT-PT}
         (row-heights example-parsed-1))))

(deftest test-cell-styles
  (let [styles (cell-styles example-parsed-no-div 8)]
    (is (= #{:bold :italic} (styles [0 0]) (styles [1 0]) (styles [2 0])))
    (is (nil? (styles [3 0])))
    (is (nil? (styles [5 0])) "the label column has no borders")
    (is (nil? (styles [5 9])) "nothing beyond the data columns")
    (testing "heading rows"
      (is (= #{:border-left :border-right :border-top :border-bottom :center :wrap}
             (styles [5 1]) (styles [6 8]))))
    (testing "end, paragraph and reference rows are rotated"
      (is (= #{:border-left :border-right :wrap :rotate-90}
             (styles [7 1]) (styles [8 1])))
      (is (= #{:border-left :border-right :border-bottom :wrap :rotate-90}
             (styles [9 1]))))))

(deftest test-hzd-layout
  (let [layout (hzd-layout example-input-no-div)]
    (is (= example-parsed-no-div (:rows layout)))
    (is (= (merge-ranges example-parsed-no-div) (:merges layout)))
    (is (= 8 (count (:col-widths layout))))
    (is (= con/DEFAULT-COL-WIDTH-CHARS (:default-col-width-chars layout)))
    (is (= (row-heights example-parsed-no-div) (:row-heights layout)))
    (is (= (cell-styles example-parsed-no-div 8) (:cell-styles layout)))))
