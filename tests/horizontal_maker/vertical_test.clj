(ns horizontal-maker.vertical-test
  (:require  [clojure.test :as t]
             [horizontal-maker.vertical :refer :all]))

(def example-input-no-div [[1 "Paragraph 1" "" "Appoint"    "Section 1" ""]
                           [5 "Para 2"      "" ""           ""           ""]
                           [16 "Para 3"     20 "Elders"     ""           ""]
                           [1 "Para 4"      ""  ""          ""           ""]
                           [15 "Para 5"     ""  "More"      "Section 2"  ""]
                           [21 "Para 6"     30  ""          ""           ""]
                           [1 "Para 7"      ""  "Great"     ""           ""]
                           [10 "Para 8"     15          ""           ""]])

(t/deftest test-generate-vertical-parsed
  (t/is (= [["Appoint" ["Paragraph 1" "Para 2"] "Romans 1:1-15"]
            ["Elders" ["Para 3" "Para 4"] "Romans 1:16-2:14"]
            ["More" ["Para 5" "Para 6"] "Romans 2:15-30"]
            ["Great" ["Para 7" "Para 8"] "Romans 3:1-15"]]
           (generate-vertical-parsed example-input-no-div "Romans"))))
