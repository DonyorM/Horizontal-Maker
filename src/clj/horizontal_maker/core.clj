(ns horizontal-maker.core
  "Pure functions that turn the rows of an input sheet into the layout of a
  horizontal. Nothing here touches a file; see horizontal-maker.excel and
  horizontal-maker.ods for the writers."
  (:require [horizontal-maker.constants :as con]))

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
  "Gets the chapter-end references for each segment (or other division)."
  [input ref-idx seg-idx]
  (loop [row (second input)
         rows (drop 2 input)
         refs []
         current-end ""]
    (let [chap-end (if-not (empty? (str (nth row ref-idx)))
                     (nth row ref-idx)
                     current-end)]
      (if (= 0 (count rows))
        (conj refs chap-end)
        (if-not (empty? (nth row seg-idx))
          (recur (first rows) (rest rows) (conj refs current-end) (nth row ref-idx))
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
                     [(str chap ":" (int (nth row ref-idx)))]
                     []))
      (recur (first rows) (rest rows) (if-not (empty? (str (nth row con/END-IDX)))
                                        (inc chap)
                                        chap)
             (if-not (empty? (nth row seg-idx))
               (conj refs (str chap ":" (int (nth row ref-idx))))
               refs)))))

(defn- division-based-sheet
  "Creates a sheet with divisions"
  [input]
  (loop [row (first input)
         rows (rest input)
         divs ["Divis."]
         sects ["Sect."]
         segms ["Segm."]]
    (let [seg (nth row con/SEG-IDX)
          seg? (seq seg)
          sect (nth row con/SECT-IDX)
          div (nth row con/DIV-IDX)
          result-vec [(concat divs (if seg? [div] []))
                      (concat sects (if seg? [sect] []))
                      (concat segms (if seg? [seg] []))]]
      (if (= 0 (count rows))
        [(first result-vec) (second result-vec)
         (cons "" (get-end-references input con/END-IDX con/SEG-IDX)) (last result-vec)
         (cons "" (get-start-references input con/START-IDX con/SEG-IDX))]
        (recur (first rows) (rest rows) (first result-vec)
               (second result-vec) (last result-vec))))))

(defn- paragraph-based-sheet
  "For a sheet without divisions."
  [input]
  (let [first-row (first input)
        ;; Transpose the input so each input column becomes a row, labelled
        ;; with the heading shown in the first column of the horizontal.
        labels ["" "" "Paragr." "" "Segm." "Sect."]
        transposed (apply map #(identity %&) labels input)]
    (reverse (cons (cons "" (get-start-references input con/START-IDX con/PARA-IDX))
                   ;; Drop the chapter and verse rows; they aren't shown.
                   (drop 2 (take (cond
                                   (seq (nth first-row con/SECT-IDX)) 6
                                   (seq (nth first-row con/SEG-IDX)) 5
                                   true 4)
                                 transposed))))))

(defn create-spreadsheet-vector
  "Parse the input organized in columns to output organized in rows
  (as presented in the final horizontal.) This output is suitable to be passed
  to add-rows!

  It expects a vector of vectors with the following schema:
  [chapter start-verse title chapter-end-verse? segment? section? division?]"
  [input]
  (let [first-row (first input)
        division? (seq (nth first-row con/DIV-IDX))]
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
  (if not present). Expects a un-parsed input.

  t-width is the total width, defaults to con/TOTAL-WIDTH"
  ([input t-width]
   (let [total-verses (apply + (remove #(empty? (str %))
                                       (map #(nth % con/END-IDX) input)))
         div? (seq (nth (first input) con/DIV-IDX))
         verse-dif (fn [s ps pv]
                     (* t-width (/ (+ (- s ps)
                                      pv) total-verses)))]
     (loop [row (second input)
            rows (drop 2 input)
            previous-start 1
            previous-value 0
            width []]
       (let [start (nth row con/START-IDX)
             end? (seq (str (nth row con/END-IDX)))
             new-seg? (seq (nth row con/SEG-IDX))]
         (if (= 0 (count rows))
           (map int (as-> width w
                      (if (or (not div?) new-seg?)
                        (conj w (verse-dif start previous-start previous-value))
                        w)
                      (conj w (* t-width (/ (+
                                             (- (nth row con/END-IDX) start)
                                             1
                                             (if (and div? (not new-seg?))
                                               (- start previous-start)
                                               0)
                                             previous-value)
                                            total-verses)))))
           (recur (first rows) (rest rows) (if end? 1 start)
                  (+
                   (if end?
                     (- (nth row con/END-IDX) start -1)
                     0)
                   (if (and div? (not new-seg?))
                     (+ (- start previous-start) previous-value)
                     0))
                  (if div?
                    (if new-seg?
                      (conj width (verse-dif start previous-start previous-value))
                      width)
                    (conj width (verse-dif start previous-start previous-value)))))))))
  ([input]
   (calculate-width input con/TOTAL-WIDTH)))

;;; Layout: everything a writer needs to know about the HZD sheet, independent
;;; of the file format. Rows and columns are 0-based.

(def ^:private HEADER-ROWS
  "Number of rows above the horizontal itself (Book, Title, Key Verse, 2 blank)."
  5)

(defn merge-ranges
  "The cells to merge in order to form the headings (divisions, sections, or
  segments), as [row-idx first-col last-col] triples. Only ranges spanning more
  than one cell are returned.

  Expects pre-parsed input, as generated by create-spreadsheet-vector"
  [parsed]
  (let [headers (->> parsed
                     (drop HEADER-ROWS)
                     (drop-last 3))] ; The last three rows (End, Para/Seg, Ref) never merge
    (vec (mapcat (fn [row-idx row]
                   (for [[c1 c2] (get-merge-range row)
                         :when (> c2 c1)]
                     [row-idx c1 c2]))
                 (range HEADER-ROWS (+ HEADER-ROWS (count headers)))
                 headers))))

(defn row-heights
  "Map of row index to height in points for the rows that aren't default
  height: the paragraph (or segment) title row and the reference row."
  [parsed]
  (let [n (count parsed)]
    {(- n 2) con/PARA-ROW-HEIGHT-PT
     (- n 1) con/REF-ROW-HEIGHT-PT}))

(defn cell-styles
  "Map of [row col] to a set of style keywords for every cell of the horizontal
  that needs formatting. n-cols is the number of data columns (excluding the
  label column).

  Keywords: :bold :italic :border-left :border-right :border-top :border-bottom
  :center :wrap :rotate-90"
  [parsed n-cols]
  (let [n (count parsed)
        body-rows (range HEADER-ROWS n)
        rotated-rows (set (take-last 3 body-rows))
        last-row (dec n)
        heading-styles (into {} (for [r (range 3)] [[r 0] #{:bold :italic}]))]
    (into heading-styles
          (for [r body-rows
                c (range 1 (inc n-cols))]
            [[r c] (cond-> #{:border-left :border-right :wrap}
                     (not (rotated-rows r)) (conj :border-top :border-bottom :center)
                     (rotated-rows r) (conj :rotate-90)
                     (= r last-row) (conj :border-bottom))]))))

(defn hzd-layout
  "Everything needed to write the HZD sheet, independent of file format:

  {:rows        parsed rows, suitable for writing cell by cell
   :merges      [[row first-col last-col] ...]
   :col-widths  [w ...] widths of the data columns (1..n) in Excel units
   :default-col-width-chars  width of other columns, in characters
   :row-heights {row-idx points}
   :cell-styles {[row col] #{style-keyword ...}}}"
  ([input t-width]
   (let [parsed (mapv vec (create-spreadsheet-vector input))
         widths (vec (calculate-width input t-width))]
     {:rows parsed
      :merges (merge-ranges parsed)
      :col-widths widths
      :default-col-width-chars con/DEFAULT-COL-WIDTH-CHARS
      :row-heights (row-heights parsed)
      :cell-styles (cell-styles parsed (count widths))}))
  ([input]
   (hzd-layout input con/TOTAL-WIDTH)))
