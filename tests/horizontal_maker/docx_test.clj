(ns horizontal-maker.docx-test
  (:require [clojure.java.io :as io]
            [clojure.test :refer :all]
            [horizontal-maker.formats :as formats]
            [horizontal-maker.test-helpers :as h]
            [horizontal-maker.vertical :as vertical])
  (:import [org.apache.poi.xwpf.usermodel XWPFDocument XWPFTable]))

(deftest test-make-vertical
  (let [path (h/temp-file "vertical.docx")
        parsed (vertical/generate-vertical-parsed h/sample-rows "Romans")]
    (formats/make-vertical! path h/sample-rows "Romans")
    (with-open [in (io/input-stream path)
                doc (XWPFDocument. in)]
      (let [tables (vec (.getTables doc))]
        (is (= (count parsed) (count tables)))
        (doseq [[[name paragraphs ref] ^XWPFTable table] (map vector parsed tables)]
          (is (= (+ 3 (count paragraphs)) (.getNumberOfRows table)))
          (is (= name (.getText (.getCell (.getRow table 0) 1))))
          (is (= ref (.getText (.getCell (.getRow table 0) 2))))
          (is (= paragraphs (map #(.getText (.getCell (.getRow table (inc %)) 1))
                                 (range (count paragraphs))))))
        (is (= "Segment 1" (.getText (.getCell (.getRow (first tables) 0) 0))))))))
