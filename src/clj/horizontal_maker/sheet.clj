(ns horizontal-maker.sheet
  "Reading input sheets and writing the HZD sheet, dispatched on the file
  extension. Implementations live in horizontal-maker.excel and
  horizontal-maker.ods."
  (:require [clojure.java.io :as io]
            [clojure.string :as str]
            [horizontal-maker.constants :as con]))

(defn extension
  "The lower-cased extension of a path as a keyword, e.g. :xlsx, or nil."
  [path]
  (let [name (.getName (io/file (str path)))
        idx (.lastIndexOf name ".")]
    (when (pos? idx)
      (keyword (str/lower-case (subs name (inc idx)))))))

(defmulti read-rows
  "Reads the first sheet of the spreadsheet at path. Returns a seq of rows,
  each a vector of exactly con/NUM-COLS values:

  [chapter verse title end segment section division]

  verse and end are Doubles (or nil when blank); every other value is a
  String (or nil when blank). Rows whose verse cell isn't a number (such as the
  header row) are dropped."
  (fn [path] (extension path)))

(defmulti write-hzd!
  "Writes the layout (see horizontal-maker.core/hzd-layout) into a sheet named
  HZD of the spreadsheet at path, replacing any existing HZD sheet. The other
  sheets are left as they are."
  (fn [path layout] (extension path)))

(defn- unsupported [path]
  (ex-info (str "Unsupported spreadsheet format: " path
                ". Use an .xls, .xlsx or .ods file.")
           {:path (str path)}))

(defmethod read-rows :default [path] (throw (unsupported path)))
(defmethod write-hzd! :default [path _] (throw (unsupported path)))

;;; Helpers shared by the implementations

(defn cell->str
  "Normalises a non-numeric cell value: nil and blank become nil, numbers that
  are whole are rendered without a decimal point."
  [v]
  (cond
    (nil? v) nil
    (and (number? v) (== v (Math/floor (double v)))) (str (long v))
    :else (let [s (str v)] (when-not (str/blank? s) s))))

(defn normalize-row
  "Pads or truncates raw cell values to con/NUM-COLS and coerces them to the
  read-rows contract."
  [raw]
  (let [cells (vec (take con/NUM-COLS (concat raw (repeat nil))))]
    (vec (map-indexed (fn [idx v]
                        (if (or (= idx con/START-IDX) (= idx con/END-IDX))
                          (when (number? v) (double v))
                          (cell->str v)))
                      cells))))

(defn data-row?
  "True when a raw row holds a paragraph: its verse cell is a number."
  [raw]
  (number? (nth raw con/START-IDX nil)))

(defn normalize-rows
  "Applies the read-rows contract to a seq of raw rows (seqs of cell values)."
  [raws]
  (->> raws
       (filter data-row?)
       (mapv normalize-row)))
