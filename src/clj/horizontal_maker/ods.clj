(ns horizontal-maker.ods
  "Reads .ods input and writes the HZD sheet with odfdom."
  (:require [clojure.java.io :as io]
            [horizontal-maker.constants :as con]
            [horizontal-maker.odf :as odf]
            [horizontal-maker.sheet :as sheet])
  (:import [java.io File]
           [java.nio.file Files StandardCopyOption]
           [org.odftoolkit.odfdom.doc OdfSpreadsheetDocument]
           [org.odftoolkit.odfdom.doc.table OdfTable OdfTableCell OdfTableRow]
           [org.odftoolkit.odfdom.dom.element.style StyleParagraphPropertiesElement
            StyleTableCellPropertiesElement StyleTextPropertiesElement]
           [org.odftoolkit.odfdom.dom.style OdfStyleFamily]))

(set! *warn-on-reflection* true)

(def ^:private MAX-BLANK-ROWS
  "Reading stops after this many consecutive blank rows. Spreadsheets saved by
  LibreOffice report a million rows, nearly all of them empty."
  50)

(def ^:private HZD "HZD")

(defn- row-values [^OdfTableRow row]
  (mapv #(odf/cell-value (.getCellByIndex row (int %))) (range con/NUM-COLS)))

(defn- table-rows
  "The raw cell values of the first con/NUM-COLS columns of each row, up to
  the last non-blank row."
  [^OdfTable table]
  (let [row-count (.getRowCount table)]
    (loop [idx 0
           blank-run 0
           rows []]
      (if (or (>= idx row-count) (>= blank-run MAX-BLANK-ROWS))
        rows
        (let [values (row-values (.getRowByIndex table idx))
              blank? (every? nil? values)]
          (recur (inc idx)
                 (if blank? (inc blank-run) 0)
                 (if blank? rows (conj rows values))))))))

(defn- first-table ^OdfTable [^OdfSpreadsheetDocument doc]
  (or (first (.getTableList doc))
      (throw (ex-info "The spreadsheet has no sheets." {}))))

(defn read-rows [path]
  (let [doc (OdfSpreadsheetDocument/loadDocument (io/file (str path)))]
    (try
      (sheet/normalize-rows (table-rows (first-table doc)))
      (finally
        (.close doc)))))

(defmethod sheet/read-rows :ods [path] (read-rows path))

(def ^:private BORDER "0.5pt solid #000000")

(defn- style-props
  "Translates a set of layout style keywords into odfdom cell style
  properties."
  [styles]
  (cond-> {}
    (:bold styles) (assoc StyleTextPropertiesElement/FontWeight "bold")
    (:italic styles) (assoc StyleTextPropertiesElement/FontStyle "italic")
    (or (:bold styles) (:italic styles)) (assoc StyleTextPropertiesElement/FontSize "11pt")
    (:border-left styles) (assoc StyleTableCellPropertiesElement/BorderLeft BORDER)
    (:border-right styles) (assoc StyleTableCellPropertiesElement/BorderRight BORDER)
    (:border-top styles) (assoc StyleTableCellPropertiesElement/BorderTop BORDER)
    (:border-bottom styles) (assoc StyleTableCellPropertiesElement/BorderBottom BORDER)
    (:center styles) (assoc StyleParagraphPropertiesElement/TextAlign "center")
    (:wrap styles) (assoc StyleTableCellPropertiesElement/WrapOption "wrap")
    (:rotate-90 styles) (assoc StyleTableCellPropertiesElement/RotationAngle "90"
                               StyleTableCellPropertiesElement/RotationAlign "bottom")))

(defn- set-cell-value! [^OdfTableCell cell value]
  (cond
    (nil? value) nil
    (number? value) (.setDoubleValue cell (Double/valueOf (double value)))
    :else (let [s (str value)]
            (when-not (= "" s)
              (.setStringValue cell s)))))

(defn- width->mm ^long [excel-width]
  (max 1 (Math/round (* (/ (double excel-width) 256) con/MM-PER-CHAR))))

(defn write-hzd!
  "Writes the layout into a new HZD sheet of the document, replacing any
  existing one. Does not save."
  [^OdfSpreadsheetDocument doc
   {:keys [rows merges col-widths default-col-width-chars row-heights cell-styles]}]
  (when-let [^OdfTable old (.getTableByName doc HZD)]
    (.remove old))
  (let [n-rows (count rows)
        n-cols (inc (count col-widths))
        table (OdfTable/newTable doc (int n-rows) (int n-cols))
        style-name (odf/style-cache doc)]
    (.setTableName table HZD)
    (doseq [[row-idx row] (map-indexed vector rows)
            [col-idx value] (map-indexed vector row)]
      (set-cell-value! (.getCellByPosition table (int col-idx) (int row-idx)) value))
    (doseq [[row-idx c1 c2] merges]
      (.merge (.getCellRangeByPosition table (int c1) (int row-idx) (int c2) (int row-idx))))
    (.setWidth (.getColumnByIndex table 0)
               (long (* default-col-width-chars con/MM-PER-CHAR)))
    (doseq [[idx width] (map-indexed vector col-widths)]
      (.setWidth (.getColumnByIndex table (int (inc idx))) (width->mm width)))
    (doseq [[row-idx points] row-heights]
      (.setHeight (.getRowByIndex table (int row-idx)) (odf/pt->mm points) false))
    (doseq [[[row-idx col-idx] styles] cell-styles]
      (let [cell (.getCellByPosition table (int col-idx) (int row-idx))]
        (odf/set-style! (.getOdfElement cell)
                        (style-name OdfStyleFamily/TableCell (style-props styles)))))
    table))

(defn save-in-place!
  "Saves the document over the file it was loaded from, by way of a temporary
  file in the same directory."
  [^OdfSpreadsheetDocument doc ^File file]
  (let [dir (or (.getParentFile (.getAbsoluteFile file)) (io/file "."))
        tmp (File/createTempFile "hzd" ".ods" dir)]
    (try
      (.save doc tmp)
      (.close doc)
      (Files/move (.toPath tmp) (.toPath file)
                  (into-array [StandardCopyOption/REPLACE_EXISTING]))
      (finally
        (.delete tmp)))))

(defmethod sheet/write-hzd! :ods [path layout]
  (let [file (io/file (str path))
        doc (OdfSpreadsheetDocument/loadDocument file)]
    (write-hzd! doc layout)
    (save-in-place! doc file)))
