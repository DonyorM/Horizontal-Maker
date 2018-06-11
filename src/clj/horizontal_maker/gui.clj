(ns horizontal-maker.gui
  (:require [horizontal-maker.core :as c]
            [horizontal-maker.vertical :as v]
            [dk.ative.docjure.spreadsheet :as sp]
            [clojure.stacktrace]
            [clojure.java.io]
            [seesaw.core :as s]
            [seesaw.chooser :as sc]
            [horizontal-maker.constants :as con]))

(defn generate [source width]
  (sp/save-workbook! source
                     (c/make-hzd source width)))

(defn display-error [message e]
  (s/show! (s/dialog
            :minimum-size [580 :by 340]
            :type :error
            :content (s/vertical-panel :items [(s/label message)
                                               (s/scrollable
                                                (s/text
                                                :text (let [writ (java.io.StringWriter.)]
                                                        (.printStackTrace e (java.io.PrintWriter. writ))
                                                        (.toString writ))
                                                :editable? false
                                                :multi-line? true))]))))

(defn show-vertical-dialog [parent input-loc]
  (let [output-field (s/text)
        browse (s/button :text "Browse"
                         :listen [:action (fn [& args]
                                            (let [file (sc/choose-file
                                                        :all-files? false
                                                        :type :save
                                                        :filters [(sc/file-filter
                                                                   "Documents"
                                                                   #(or (.endsWith (.getName %) ".docx")
                                                                        (.isDirectory %)))])]
                                            (s/text! output-field (cond-> (.getPath file)
                                                                    (not (.endsWith (.getPath file) ".docx")) (str ".docx")))))])
        book-field (s/text)
        generate (s/button :text "Generate"
                           :listen [:action (fn [& args]
                                              (try
                                                (let [file (clojure.java.io/file (s/text output-field))]
                                                  (if (or (not (.exists file))
                                                        (= :success
                                                           (s/show!
                                                            (s/dialog
                                                             :minimum-size [400 :by 130]
                                                             :content (str (.getName file) " already exists. Overwrite?")
                                                             :type :question
                                                             :option-type :yes-no))))
                                                  (do
                                                    (v/make-vertical
                                                     (s/text output-field)
                                                     (c/get-values (sp/load-workbook input-loc))
                                                     (s/text book-field))
                                                    (s/alert "Vertical created."))))
                                                (catch Exception e
                                                  (clojure.stacktrace/print-stack-trace e)
                                                  (display-error "Error generating vertical. Please have someone knowledgeable look at the logs." e))))])
        close (s/button :text "Back")
        frame (s/frame :title "Make Vertical"
                       :minimum-size [490 :by 90]
                       :content (s/vertical-panel
                                 :items [(s/horizontal-panel
                                          :items [(s/label "Save to:")
                                                  output-field
                                                  browse])
                                         (s/horizontal-panel
                                          :items [(s/label "Book Name:") book-field])
                                         (s/horizontal-panel
                                          :items [generate close])]))]
    (s/listen close :action (fn [& args]
                               (s/hide! frame)
                               (s/show! parent)))
    (-> frame s/pack! s/show!)))

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
        width-field (s/spinner :model con/TOTAL-WIDTH)
        done (s/button :text "Generate Horizontal"
                       :listen [:action (fn [& args] (try
                                                       (generate (s/text input-field)
                                                                 (s/selection width-field))
                                           (s/alert "Horizontal Created. Please close and re-open your excel sheet to see the changes.")
                                           (catch Exception e
                                             (clojure.stacktrace/print-stack-trace e)
                                             (display-error "Error generating horizontal. Ask someone knowledgeable to see the logs." e))))])
        vertical (s/button :text "Generate Vertical")
        frame (s/frame :title "Make Horizontal"
                       :minimum-size [490 :by 90]
                       :on-close :dispose
                       :content (s/vertical-panel :items
                                                  [(s/horizontal-panel
                                                    :items [(s/label "Input File:")
                                                            input-field browse])
                                                   (s/horizontal-panel
                                                    :items [(s/label "Total Width:")
                                                            width-field])
                                                   (s/horizontal-panel
                                                    :items [done vertical])]))]
    (s/listen vertical :action (fn [& args]
                                 (s/hide! frame)
                                 (show-vertical-dialog frame (s/text input-field))))
    (-> frame s/pack! s/show!)))
