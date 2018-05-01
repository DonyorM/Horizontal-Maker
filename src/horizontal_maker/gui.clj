(ns horizontal-maker.gui
  (:require [horizontal-maker.core :as c]
            [dk.ative.docjure.spreadsheet :as sp]
            [clojure.stacktrace]
            [seesaw.core :as s]
            [seesaw.chooser :as sc]))

(defn generate [source]
  (sp/save-workbook! source
                     (c/make-hzd source)))

(defn start-gui []
  (s/native!)
  (let [input-field (s/text)
        browse (s/button :text "Browse"
                         :listen [:action #(let [t %
                                                 f (sc/choose-file
                                                    :all-files? false
                                                    :filters [(sc/file-filter "Excel Files" (fn [x] (or
                                                                                                     (.endsWith (.getName x) ".xls")

                                                                                                     (.endsWith (.getName x) ".xlsx")
                                                                                                     (.isDirectory x))))
                                                              (sc/file-filter "All Files" (fn [& args] (constantly true)))])]
                                             (s/text! input-field (if f
                                                                    (.getPath f)
                                                                    "")))])
        done (s/button :text "Generate"
                       :listen [:action (fn [& args] (try
                                           (generate (s/text input-field))
                                           (s/alert "Horizontal Created. Please close and re-open your excel sheet to see the changes.")
                                           (catch Exception e
                                             (clojure.stacktrace/print-stack-trace e)
                                             (s/alert "Error generating horizontal. Ask someone knowledgeable to see the logs.")
                                             (throw e))))])
        frame (s/frame :title "Make Horizontal"
                       :minimum-size [490 :by 90]
                       :content (s/vertical-panel :items
                                                  [(s/horizontal-panel
                                                    :items [(s/label "Input File:")
                                                            input-field browse])
                                                   done]))]
    (-> frame s/pack! s/show!)))
