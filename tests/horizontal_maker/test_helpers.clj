(ns horizontal-maker.test-helpers
  (:require [clojure.java.io :as io])
  (:import [java.nio.file Files]
           [java.nio.file.attribute FileAttribute]))

(def sample-sheet
  "The README example, as rows of raw cell values, with a header row and a
  blank Chapter cell on most rows."
  [["Chapter" "Verse" "Paragraph Title" "End" "Segment" "Section" "Division"]
   ["1" 1 "Four words only" nil "Segment 1" "Section 1" nil]
   [nil 5 "Paragraph 2" nil nil nil nil]
   [nil 10 "Paragraph 3" 15 "Segment 2" nil nil]
   ["2" 1 "Paragraph 4" nil nil nil nil]
   [nil 13 "Paragraph 5" 17 nil nil nil]
   ["3" 2 "Paragraph 6" nil "Segment 3" "Section 2" nil]
   [nil 9 "Paragraph 7" nil nil nil nil]
   [nil 15 "Paragraph 8" 25 "Segment 4" nil nil]
   ["4" 1 "Paragraph 9" nil nil nil nil]
   [nil 12 "Paragraph 10" 22 nil nil nil]])

(def sample-rows
  "What read-rows should return for sample-sheet."
  [["1" 1.0 "Four words only" nil "Segment 1" "Section 1" nil]
   [nil 5.0 "Paragraph 2" nil nil nil nil]
   [nil 10.0 "Paragraph 3" 15.0 "Segment 2" nil nil]
   ["2" 1.0 "Paragraph 4" nil nil nil nil]
   [nil 13.0 "Paragraph 5" 17.0 nil nil nil]
   ["3" 2.0 "Paragraph 6" nil "Segment 3" "Section 2" nil]
   [nil 9.0 "Paragraph 7" nil nil nil nil]
   [nil 15.0 "Paragraph 8" 25.0 "Segment 4" nil nil]
   ["4" 1.0 "Paragraph 9" nil nil nil nil]
   [nil 12.0 "Paragraph 10" 22.0 nil nil nil]])

(defn temp-dir
  "A fresh temporary directory, as a File."
  []
  (.toFile (Files/createTempDirectory "horizontal-maker" (make-array FileAttribute 0))))

(defn temp-file
  "A path inside a fresh temporary directory."
  [name]
  (.getPath (io/file (temp-dir) name)))
