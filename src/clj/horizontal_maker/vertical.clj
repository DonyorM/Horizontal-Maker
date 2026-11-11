(ns horizontal-maker.vertical
  "Builds the vertical (one page per segment listing its paragraphs) from the
  input rows. Writing is dispatched on the output file's extension; the .docx
  writer is the Java class horizontal_maker.Vertical, the .odt writer lives in
  horizontal-maker.odt."
  (:require [horizontal-maker.constants :as con]
            [horizontal-maker.sheet :as sheet])
  (:import horizontal_maker.Vertical))

(defn generate-vertical-parsed
  "Returns one entry per segment: [segment-name [paragraph-strings] reference].
  Paragraph strings are \"chapter:verse Title\"; the reference is
  \"Book c:v-c:v\"."
  [input book]
  (let [find-end (fn [start-chap start-verse chap end-verse]
                   (str book " " start-chap ":" (int start-verse) "-"
                        (if (not= start-chap chap)
                          (str chap ":"))
                        (int end-verse)))]
    (loop [row (second input)
           rows (drop 2 input)
           result []
           seg-name (nth (first input) con/SEG-IDX)
           paragraphs [(str 1 ":" (int (nth (first input) con/START-IDX)) " "
                            (nth (first input) con/PARA-IDX))]
           chap 1
           start-chap 1
           start-verse 1
           prev-end ""]
      (if (= 0 (count rows))
        (conj result [seg-name (conj paragraphs (str chap ":" (int (nth row con/START-IDX))
                                                     " " (nth row con/PARA-IDX)))
                      (find-end start-chap start-verse chap (nth row con/END-IDX))])
        (let [new-chap (if (empty? (str (nth row con/END-IDX)))
                         chap
                         (inc chap))
              new-seg? (seq (nth row con/SEG-IDX))]
          (recur (first rows)
                 (rest rows)
                 (if (empty? (nth row con/SEG-IDX))
                   result
                   (conj result [seg-name paragraphs
                                 (find-end start-chap start-verse
                                           (if (empty? (str prev-end))
                                             chap
                                             (dec chap))
                                           (if (empty? (str prev-end))
                                             (dec (nth row con/START-IDX))
                                             prev-end))]))
                 (if new-seg?
                   (nth row con/SEG-IDX)
                   seg-name)
                 (conj (if new-seg?
                         []
                         paragraphs)
                       (str chap ":" (int (nth row con/START-IDX)) " " (nth row con/PARA-IDX)))
                 new-chap
                 (if new-seg?
                   chap
                   start-chap)
                 (if new-seg?
                   (nth row con/START-IDX)
                   start-verse)
                 (nth row con/END-IDX)))))))

(defmulti write-vertical!
  "Writes the vertical to the output path. paragraphs is a seq of seqs of
  paragraph strings, names a seq of segment names and refs a seq of
  references, all one per segment."
  (fn [output paragraphs names refs] (sheet/extension output)))

(defmethod write-vertical! :docx [output paragraphs names refs]
  (Vertical/createDocument (str output) paragraphs names refs))

(defmethod write-vertical! :default [output & _]
  (throw (ex-info (str "Unsupported document format: " output
                       ". Use a .docx or .odt file.")
                  {:path (str output)})))

(defn make-vertical
  "Generates the vertical for the segments start (inclusive) to end
  (exclusive) and writes it to output."
  ([output input book-name start end]
   (let [parsed (subvec (vec (generate-vertical-parsed input book-name)) start end)]
     (write-vertical! output (map second parsed) (map first parsed) (map last parsed))))
  ([output input book-name]
   (make-vertical output input book-name 0
                  (count (remove #(empty? (nth % con/SEG-IDX)) input)))))
