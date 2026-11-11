(ns horizontal-maker.formats
  "Entry points for generating horizontals and verticals from a file of any
  supported format. Requiring this namespace registers every reader and
  writer."
  (:require [horizontal-maker.constants :as con]
            [horizontal-maker.core :as core]
            [horizontal-maker.excel]
            [horizontal-maker.ods]
            [horizontal-maker.odt]
            [horizontal-maker.sheet :as sheet]
            [horizontal-maker.vertical :as vertical]))

(def spreadsheet-extensions #{:xls :xlsx :ods})
(def document-extensions #{:docx :odt})

(defn read-input
  "Reads the input rows from a spreadsheet. See horizontal-maker.sheet/read-rows."
  [path]
  (sheet/read-rows path))

(defn make-hzd!
  "Generates the horizontal from the first sheet of the spreadsheet at path
  and writes it into that file's HZD sheet. t-width is the total width in
  Excel column width units."
  ([path t-width]
   (sheet/write-hzd! path (core/hzd-layout (read-input path) t-width)))
  ([path]
   (make-hzd! path con/TOTAL-WIDTH)))

(defn make-vertical!
  "Generates the vertical for the input rows and writes it to output (.docx
  or .odt)."
  [output input book-name]
  (vertical/make-vertical output input book-name))
