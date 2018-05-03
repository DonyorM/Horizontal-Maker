(ns horizontal-maker.vertical
  (:require [horizontal-maker.constants :as con])
  (:import horizontal_maker.Vertical))

(defn generate-vertical-parsed
  "Converts a input format for a horizontal into one suitable for a vertical."
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
                      (find-end start-chap start-verse
                                chap (nth row con/END-IDX))])
        ;;else
        (let [new-chap (if (empty? (str (nth row con/END-IDX)))
                         chap
                         (inc chap))
              new-seg? (seq (nth row con/SEG-IDX))]
          (recur (first rows) (rest rows) (if (empty? (nth row con/SEG-IDX))
                                            result
                                            (conj result [seg-name paragraphs
                                                          (find-end start-chap start-verse (if (empty? (str prev-end)) chap (dec chap))
                                                                    (if (empty? (str prev-end))
                                                                      (dec (nth row con/START-IDX))
                                                                      prev-end))]))
                 (if new-seg?
                   (nth row con/SEG-IDX)
                   seg-name)
                 (conj (if new-seg? [] paragraphs) (str chap ":"
                                                        (int (nth row con/START-IDX))
                                                        " " (nth row con/PARA-IDX)))
                 new-chap
                 (if new-seg?
                   chap
                   start-chap)
                 (if new-seg?
                   (nth row con/START-IDX)
                   start-verse)
                 (nth row con/END-IDX)))))))

(defn make-vertical
  "Generates a .docx vertical file and saves it to output. If start and end not passed, it will default to every segment.

  start: the first segment to add to the vertical (as indexed in the input)
  end: the last segment (not inclusive) to add to the vertical."
  ([output input book-name start end]
   (let [parsed (subvec (vec (generate-vertical-parsed input book-name)) start end)]
     (Vertical/createDocument output (map second parsed) (map first parsed)
                              (map last parsed))))
  ([output input book-name]
   (make-vertical output input book-name 0 (count (remove #(empty? (nth % con/SEG-IDX)) input)))))
