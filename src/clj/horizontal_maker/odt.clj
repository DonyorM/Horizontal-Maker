(ns horizontal-maker.odt
  "Writes the vertical as an OpenDocument text file (.odt) with odfdom. The
  layout mirrors horizontal_maker.Vertical, the .docx writer."
  (:require [clojure.java.io :as io]
            [horizontal-maker.odf :as odf]
            [horizontal-maker.vertical :as vertical])
  (:import [org.odftoolkit.odfdom.doc OdfTextDocument]
           [org.odftoolkit.odfdom.doc.table OdfTable OdfTableCell]
           [org.odftoolkit.odfdom.dom.element.style StyleParagraphPropertiesElement
            StyleTextPropertiesElement]
           [org.odftoolkit.odfdom.dom.element.text TextPElement]
           [org.odftoolkit.odfdom.dom.style OdfStyleFamily]))

(set! *warn-on-reflection* true)

(def TABLE-HEIGHT-MM "Height of the rows below the header, together." 194)
(def COL-WIDTH-MM "Width of each of the three columns." 53)
(def HEADER-FONT-SIZE "14pt")
(def SEGMENT-COLOR "#007171")
(def APPLICATION-COLOR "#007826")

(defn- paragraph-props
  "odfdom paragraph style properties for the text options. The alignment is
  one of \"start\", \"center\" or \"end\"."
  [{:keys [align bold underline color size]}]
  (cond-> {}
    align (assoc StyleParagraphPropertiesElement/TextAlign align)
    bold (assoc StyleTextPropertiesElement/FontWeight "bold")
    underline (assoc StyleTextPropertiesElement/TextUnderlineStyle "solid"
                     StyleTextPropertiesElement/TextUnderlineWidth "auto"
                     StyleTextPropertiesElement/TextUnderlineColor "font-color")
    color (assoc StyleTextPropertiesElement/Color color)
    size (assoc StyleTextPropertiesElement/FontSize size)))

(defn- set-text!
  "Sets the cell's text and, when options are given, styles its paragraph."
  [style-name ^OdfTableCell cell ^String text options]
  (.setStringValue cell text)
  (when (seq options)
    (when-let [^TextPElement p (first (odf/child-elements (.getOdfElement cell) TextPElement))]
      (odf/set-style! p (style-name OdfStyleFamily/Paragraph (paragraph-props options))))))

(defn- add-segment!
  "Adds the table for one segment to the document."
  [^OdfTextDocument doc style-name paragraphs ^String segment-name segment-count ^String ref]
  (let [para-count (count paragraphs)
        ;; One extra row for the header, one for the summary and one for the application
        n-rows (+ para-count 3)
        table (OdfTable/newTable doc (int n-rows) 3)
        summary-row (inc para-count)
        application-row (+ para-count 2)
        text! (partial set-text! style-name)]
    (doseq [i (range 3)]
      (.setWidth (.getColumnByIndex table (int i)) COL-WIDTH-MM))
    (text! (.getCellByPosition table 0 0) (str "Segment " segment-count)
           {:align "start" :color SEGMENT-COLOR :size HEADER-FONT-SIZE})
    (text! (.getCellByPosition table 1 0) segment-name {:align "center" :size HEADER-FONT-SIZE})
    (text! (.getCellByPosition table 2 0) ref {:align "end" :size HEADER-FONT-SIZE})
    (let [row-height (quot TABLE-HEIGHT-MM (dec n-rows))]
      (doseq [i (range 1 n-rows)]
        (.setHeight (.getRowByIndex table (int i)) row-height true)))
    (doseq [[i title] (map-indexed vector paragraphs)]
      (text! (.getCellByPosition table 1 (int (inc i))) (str title) {:underline true}))
    (text! (.getCellByPosition table 1 summary-row) "Summary:" {:bold true})
    (text! (.getCellByPosition table 0 application-row) "Application:" {:color APPLICATION-COLOR})
    ;; Don't merge the header row, so start at 1
    (.merge (.getCellRangeByPosition table 0 1 0 summary-row))
    (.merge (.getCellRangeByPosition table 2 1 2 summary-row))
    ;; Merge the application row (last row)
    (.merge (.getCellRangeByPosition table 0 application-row 2 application-row))
    table))

(defn- add-page-break! [^OdfTextDocument doc style-name]
  (let [p (.newTextPElement (.getContentRoot doc))]
    (odf/set-style! p (style-name OdfStyleFamily/Paragraph
                                  {StyleParagraphPropertiesElement/BreakBefore "page"}))))

(defn- remove-empty-paragraphs!
  "A new text document starts with an empty paragraph; drop it so the first
  table sits at the top of the page."
  [^OdfTextDocument doc]
  (let [root (.getContentRoot doc)]
    (doseq [^TextPElement p (odf/child-elements root TextPElement)
            :when (= "" (.getTextContent p))]
      (.removeChild root p))))

(defn create-document
  "Writes a vertical with one table per segment to the output path."
  [output paragraphs names refs]
  (let [doc (OdfTextDocument/newTextDocument)
        style-name (odf/style-cache doc)]
    (try
      (remove-empty-paragraphs! doc)
      (doseq [[i paras name ref] (map vector (range) paragraphs names refs)]
        (add-segment! doc style-name paras name (inc i) ref)
        (add-page-break! doc style-name))
      (.save doc (io/file (str output)))
      (finally
        (.close doc)))))

(defmethod vertical/write-vertical! :odt [output paragraphs names refs]
  (create-document output paragraphs names refs))
