(ns horizontal-maker.vertical-test
  (:require [clojure.test :as t]
            [horizontal-maker.vertical :refer :all]))

(def example-input-no-div [["1" 1 "Paragraph 1" "" "Appoint"    "Section 1" ""]
                           [""  5 "Para 2"      "" ""           ""           ""]
                           [""  16 "Para 3"     20 "Elders"     ""           ""]
                           ["2" 1 "Para 4"      ""  ""          ""           ""]
                           [""  15 "Para 5"     ""  "More"      "Section 2"  ""]
                           [""  21 "Para 6"     30  ""          ""           ""]
                           ["3" 1 "Para 7"      ""  "Great"     ""           ""]
                           [""  10 "Para 8"     15  ""          ""           ""]])

(def example-parsed [["Appoint" ["1:1 Paragraph 1" "1:5 Para 2"] "Romans 1:1-15"]
                     ["Elders" ["1:16 Para 3" "2:1 Para 4"] "Romans 1:16-2:14"]
                     ["More" ["2:15 Para 5" "2:21 Para 6"] "Romans 2:15-30"]
                     ["Great" ["3:1 Para 7" "3:10 Para 8"] "Romans 3:1-15"]])

(t/deftest test-generate-vertical-parsed
  (t/is (= example-parsed
           (generate-vertical-parsed example-input-no-div "Romans"))))

(t/deftest test-write-vertical-dispatch
  (t/is (thrown-with-msg? clojure.lang.ExceptionInfo #"Unsupported document format"
                          (make-vertical "out.pdf" example-input-no-div "Romans"))))
