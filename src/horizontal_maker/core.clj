(ns horizontal-maker.core
  (:require [dk.ative.docjure.spreadsheet :as sp])
  (:import org.apache.poi.ss.util.CellRangeAddress))

;;The indexes in the input for each item of the input. See create-spreadsheet-vector docs.
(def START-IDX 0)
(def PARA-IDX 1)
(def END-IDX 2)
(def SEG-IDX 3)
(def SECT-IDX 4)
(def DIV-IDX 5)
(def TOTAL-WIDTH "The total width of the horizontal" 60000)

(defn cell-range-address [rowStart rowEnd columnStart columnEnd]
  (CellRangeAddress. rowStart rowEnd columnStart columnEnd))

(defn get-column
  "Rotates a column to a row."
  [input index]
  (reduce #(conj %1 (nth %2 index)) [] input))

(defn get-merge-range
  "Gets the range of cells that should be merged in order to form a heading for
  the passed parsed row."
  [row]
  (loop [val (nth row 2)
         row (drop 3 row)
         idx 2 #_"The first column is reserved for the description of the row"
         start_idx 1
         indexes []]
    (if (= 0 (count row))
      (concat indexes [[start_idx idx]])
      (if-not (empty? (str val))
        (recur (first row) (rest row) (inc idx) idx
               (concat indexes [[start_idx (dec idx)]]))
        (recur (first row) (rest row) (inc idx) start_idx indexes)))))

(defn get-end-references
  "Gets the references (start or end) for a segment or other divsion."
  [input ref-idx seg-idx]
  (loop [row (second input)
         rows (drop 2 input)
         refs []
         current-end ""]
    (let [chap-end (if-not (empty? (str (nth row 2)))
                     (nth row 2)
                     current-end)]
    (if (= 0 (count rows))
      (conj refs chap-end)
      (if-not (empty? (nth row seg-idx))
        (recur (first rows) (rest rows) (conj refs current-end) (nth row 2))
        (recur (first rows) (rest rows) refs chap-end))))))

(defn get-start-references
  "Gets the chapter and verse references for starting a segment"
  [input ref-idx seg-idx]
  (loop [row (first input)
         rows (rest input)
         chap 1
         refs []]
    (if (= 0 (count rows))
      (concat refs (if-not (empty? (nth row seg-idx))
                     [(str chap ":" (nth row ref-idx))]
                     []))
      (recur (first rows) (rest rows) (if-not (empty? (str (nth row END-IDX)))
                                        (inc chap)
                                        chap)
             (if-not (empty? (nth row seg-idx))
               (conj refs (str chap ":" (nth row ref-idx)))
               refs)))))

(defn- division-based-sheet
  "Creates a sheet with divisions"
  [input]
  (loop [row (first input)
         rows (rest input)
         divs ["Divis."]
         sects ["Sect."]
         segms ["Segm."]]
    (let [seg (nth row SEG-IDX)
          seg? (seq seg)
          sect (nth row SECT-IDX)
          div (nth row DIV-IDX)
          result-vec [(concat divs (if seg? [div] []))
                      (concat sects (if seg? [sect] []))
                      (concat segms (if seg? [seg] []))]]
    (if (= 0 (count rows))
      [(first result-vec) (second result-vec)
       (cons "" (get-end-references input END-IDX SEG-IDX)) (last result-vec)
       (cons "" (get-start-references input START-IDX SEG-IDX))]
      (recur (first rows) (rest rows) (first result-vec)
             (second result-vec) (last result-vec))))))

(defn- paragraph-based-sheet
  "For a sheet without divisions."
  [input]
  (let [first-row (first input)]
    (reverse (cons (cons "" (get-start-references input START-IDX PARA-IDX))
                   (rest (take (cond
                           (seq (nth first-row SECT-IDX)) 5
                           (seq (nth first-row SEG-IDX)) 4
                           true 3)
                         (apply map #(identity %&) ["" "Paragr." "" "Segm." "Sect."] input)))))))

(defn create-spreadsheet-vector
  "Parse the input organized in columns to output organized in rows
  (as presented in the final horizontal.) This output is suitable to be passed
  to add-rows!

  It expects a vector of vectors with the following schema:
  [start-verse title chapter-end-verse? segment? section? division?]"
  [input]
  (let [first-row (first input)
        division? (seq (last first-row))
        section? (seq (nth first-row SECT-IDX))
        segment? (seq (nth first-row SEG-IDX))]
    (concat
     [["Book:"]
     ["Title:"]
     ["Key Verse:"]
     [""]
     [""]]
     (if division?
       (division-based-sheet input)
       (paragraph-based-sheet input)))))

(defn calculate-width
  "Calculates the width for each segment (if divisons present), or paragaraph
  (if not present). Expects a un-parsed input."
  [input]
  (let [total-verses (apply + (remove #(empty? (str %))
                                      (map #(nth % END-IDX) input)))
        div? (seq (nth (first input) DIV-IDX))
        verse-dif (fn [s ps pv]
                    (* TOTAL-WIDTH (/ (+ (- s ps)
                                           pv) total-verses)))]
    (loop [row (second input)
           rows (drop 2 input)
           previous-start 1
           previous-value 0
           width []]
      (let [start (nth row START-IDX)
            end? (seq (str (nth row END-IDX)))
            new-seg? (seq (nth row SEG-IDX))]
        (if (= 0 (count rows))
          (map int (as-> width w
                     (if (or (not div?) new-seg?)
                       (conj w (verse-dif start previous-start previous-value))
                       w)
                     (conj w (* TOTAL-WIDTH (/ (+
                                                (- (nth row END-IDX) start)
                                                1
                                                (if (and div? (not new-seg?))
                                                  (- start previous-start)
                                                  0)
                                                previous-value)
                                               total-verses)))))
          (recur (first rows) (rest rows) (if end? 1 start)
                 (+
                  (if end?
                    (- (nth row END-IDX) start -1)
                    0)
                  (if (and div? (not new-seg?))
                    (+ (- start previous-start) previous-value)
                    0))
                 (if div?
                   (if new-seg?
                     (conj width (verse-dif start previous-start previous-value))
                     width)
                   (conj width (verse-dif start previous-start previous-value)))))))))


(defn char-range
  "Generates a range of address for columns of an excel sheet, starting from the passed character and to the offset. It will automatically increment the prefix up to ZZ (so far does not support triple letter column indexes.)

  Generally one should not pass the prefix, and allow the function to handle this itself."
  ([start offset prefix]
  ;; 90 is the value for \Z
  (if (<= (+ (int start) offset) 90)
    (map #(str prefix (char %)) (range (int start) (inc (+ offset (int start)))))
    (concat (char-range start (- 90 (int start)) prefix)
            (flatten (char-range \A
                                 (- offset (- 90 (int start)))
                                 (if (string? prefix)
                                   \A
                                   (char (inc (int prefix)))))))))
  ([start offset]
   (char-range start offset "")))

(defn merge-cells!
  "Merges the header cells (divisions, sections, or segments).

  Expects pre-parsed input, as generated by create-spreadsheet-vector"
  [parsed wb]
  (let [headers (->> parsed
                    (drop 5) ; First five rows aren't headers
                    (drop-last 3) #_"Last five never merge")]
    (doall (map (fn [row-id cell-ranges]
                  (doall (map #(if (> (- (second %) (first %)) 0)
                          (.addMergedRegion (sp/select-sheet "HZD" wb)
                                          (cell-range-address row-id row-id (first %) (second %))))
                       cell-ranges)))
                              (range 5 (+ 5 (count headers)))
                              (map get-merge-range headers)))))

(defn apply-formatting!
  "Formats each row and cell for the horizontal. Not remotely pure."
  [input wb]
  (let [PARA-HEIGHT 2400 ; Height for the main paragraph row (paragraphs or segments)
        REF-HEIGHT 540 ; Height for the row iwth starting verse reference
        DEFAULT-WIDTH 6 #_"In characters"
        sheet (sp/select-sheet "HZD" wb)
        widths (calculate-width input)
        div? (seq (nth (first input) DIV-IDX))
        sect? (seq (nth (first input) SECT-IDX))
        seg? (seq (nth (first input) SEG-IDX))
        columns-used (char-range \B (if div?
                                      (count (remove empty? (map #(nth % SEG-IDX) input)))
                                      (count input)))]
    (.setHeight (nth (sp/row-seq sheet) (cond-> 6
                                          seg? inc
                                          sect? inc)) PARA-HEIGHT)
    (.setHeight (nth (sp/row-seq sheet) (cond-> 7
                                          seg? inc
                                          sect? inc)) REF-HEIGHT)
    (.setDefaultColumnWidth sheet DEFAULT-WIDTH)
    ;; Set the formatting for the headers at the top of the page
    (doall (->> (map #(sp/select-cell % sheet) ["A1" "A2" "A3"])
            (map #(sp/set-cell-style! % (sp/create-cell-style! wb {:font {:bold true :italic true :size 11}})))))
    (doall (map (fn [index width]
                  (.setColumnWidth sheet index width))
                (range 1 (inc (count widths))) widths))
    ;; Setting borders for each cell
    (doseq [l columns-used
            n (range 6 (cond-> 9
                         seg? inc
                         sect? inc))]
      (if-let [cell (sp/select-cell (str l n) sheet)]
        (sp/set-cell-style! cell (sp/create-cell-style! wb
                                                        (cond-> {:border-left :thin
                                                                 :border-right :thin :wrap true}
                                                          (or (and sect? (< n 8))
                                                              (and (not sect?) (< n 7)))
                                                          (assoc :border-top :thin
                                                                 :border-bottom :thin
                                                                 :halign :center)
                                                          (or (>= n 10)
                                                              (and (not sect?)
                                                                   (>= n 9)))
                                                          (assoc :border-bottom :thin))))))
    ; Make text vertical for paragraph titles (or segments)
    (doseq [l columns-used
            n (range (if sect? 8 7) (if sect? 11 10))]
            (if-let [cell (sp/select-cell (str l n) sheet)]
              (.setRotation (.getCellStyle cell) 90)))))

(defn get-values
  "Reads the source from the passed workbook file"
  [source]
  (let [sheet (first (sp/sheet-seq source))]
    (map (fn [row] (map sp/read-cell (take 6 (concat row (repeat nil))))) ; The format is six columns across
         (remove #(let [val (sp/read-cell (first (sp/cell-seq %)))]
                    (or (nil? val) (not (number? val)))) (sp/row-seq sheet)))))

(defn make-hzd [source-file]
  (let [wb (sp/load-workbook source-file)
        input (get-values wb)
        parsed (create-spreadsheet-vector input)
        sheet (or (sp/select-sheet "HZD" wb)
                  (sp/add-sheet! wb "HZD"))]
    (sp/remove-all-rows! sheet)
    (.removeMergedRegions sheet (range 0 (.getNumMergedRegions sheet)))
    (sp/add-rows! sheet parsed)
    (merge-cells! parsed wb)
    (apply-formatting! input wb)
    wb))

