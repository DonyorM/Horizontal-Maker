(ns horizontal-maker.excel
  "Reads .xls/.xlsx input and writes the HZD sheet with Apache POI (docjure)."
  (:require [dk.ative.docjure.spreadsheet :as sp]
            [horizontal-maker.constants :as con]
            [horizontal-maker.sheet :as sheet])
  (:import [org.apache.poi.ss.usermodel Cell CellStyle Row Row$MissingCellPolicy Sheet Workbook]
           [org.apache.poi.ss.util CellRangeAddress]))

(set! *warn-on-reflection* true)

(defn cell-range-address [row-start row-end column-start column-end]
  (CellRangeAddress. row-start row-end column-start column-end))

(defn- row-values
  "Reads the first con/NUM-COLS cells of a row by position, so blank cells
  that were never created don't shift the later columns."
  [^Row row]
  (mapv (fn [idx]
          (some-> (.getCell row (int idx) Row$MissingCellPolicy/RETURN_BLANK_AS_NULL)
                  sp/read-cell))
        (range con/NUM-COLS)))

(defn read-workbook-rows
  "Reads the input rows (see horizontal-maker.sheet/read-rows) from the first
  sheet of a loaded workbook."
  [^Workbook wb]
  (let [sheet (first (sp/sheet-seq wb))]
    (sheet/normalize-rows (map row-values (sp/row-seq sheet)))))

(defn read-rows [path]
  (read-workbook-rows (sp/load-workbook (str path))))

(defmethod sheet/read-rows :xlsx [path] (read-rows path))
(defmethod sheet/read-rows :xls [path] (read-rows path))

(defn- make-style
  "Creates a POI cell style from a set of style keywords."
  [^Workbook wb styles]
  (let [^CellStyle style
        (sp/create-cell-style! wb (cond-> {}
                                    (or (:bold styles) (:italic styles))
                                    (assoc :font {:bold (boolean (:bold styles))
                                                  :italic (boolean (:italic styles))
                                                  :size 11})
                                    (:border-left styles) (assoc :border-left :thin)
                                    (:border-right styles) (assoc :border-right :thin)
                                    (:border-top styles) (assoc :border-top :thin)
                                    (:border-bottom styles) (assoc :border-bottom :thin)
                                    (:center styles) (assoc :halign :center)
                                    (:wrap styles) (assoc :wrap true)))]
    (when (:rotate-90 styles)
      (.setRotation style (short 90)))
    style))

(defn write-hzd!
  "Writes the layout into the HZD sheet of the workbook. Does not save."
  [^Workbook wb {:keys [rows merges col-widths default-col-width-chars row-heights cell-styles]}]
  (let [^Sheet sheet (or (sp/select-sheet "HZD" wb)
                         (sp/add-sheet! wb "HZD"))]
    (sp/remove-all-rows! sheet)
    (.removeMergedRegions sheet (range 0 (.getNumMergedRegions sheet)))
    (sp/add-rows! sheet rows)
    (doseq [[row-idx c1 c2] merges]
      (.addMergedRegion sheet (cell-range-address row-idx row-idx c1 c2)))
    (.setDefaultColumnWidth sheet (int default-col-width-chars))
    (doseq [[idx width] (map-indexed vector col-widths)]
      (.setColumnWidth sheet (int (inc idx)) (int width)))
    (doseq [[row-idx points] row-heights]
      (.setHeight ^Row (.getRow sheet (int row-idx)) (short (* 20 points))))
    (let [style-cache (atom {})
          style-for (fn [styles]
                      (or (get @style-cache styles)
                          (let [s (make-style wb styles)]
                            (swap! style-cache assoc styles s)
                            s)))]
      (doseq [[[row-idx col-idx] styles] cell-styles]
        (when-let [^Cell cell (some-> (.getRow sheet (int row-idx))
                                      (.getCell (int col-idx)))]
          (.setCellStyle cell (style-for styles)))))
    wb))

(defn- write-hzd-file! [path layout]
  (let [path (str path)
        wb (sp/load-workbook path)]
    (write-hzd! wb layout)
    (sp/save-workbook! path wb)))

(defmethod sheet/write-hzd! :xlsx [path layout] (write-hzd-file! path layout))
(defmethod sheet/write-hzd! :xls [path layout] (write-hzd-file! path layout))
