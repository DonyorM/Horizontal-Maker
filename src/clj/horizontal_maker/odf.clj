(ns horizontal-maker.odf
  "Helpers shared by the OpenDocument readers and writers (odfdom)."
  (:import [org.odftoolkit.odfdom.doc OdfDocument]
           [org.odftoolkit.odfdom.doc.table OdfTableCell]
           [org.odftoolkit.odfdom.dom.style OdfStyleFamily]
           [org.odftoolkit.odfdom.dom.style.props OdfStyleProperty]
           [org.odftoolkit.odfdom.incubator.doc.style OdfStyle]
           [org.odftoolkit.odfdom.dom.element OdfStylableElement]
           [org.w3c.dom Node]))

(set! *warn-on-reflection* true)

(defn new-style
  "Creates an automatic style of the given family in the document's content
  and sets the properties, a map of OdfStyleProperty to String. Returns the
  style's name."
  ^String [^OdfDocument doc ^OdfStyleFamily family props]
  (let [^OdfStyle style (.newStyle (.getAutomaticStyles (.getContentDom doc)) family)]
    (doseq [[^OdfStyleProperty prop ^String value] props]
      (.setProperty style prop value))
    (.getStyleNameAttribute style)))

(defn style-cache
  "Returns a function of (family props) that creates each distinct style
  once and returns its name."
  [doc]
  (let [cache (atom {})]
    (fn [family props]
      (or (get @cache [family props])
          (let [name (new-style doc family props)]
            (swap! cache assoc [family props] name)
            name)))))

(defn set-style!
  "Applies the named style to an element."
  [^OdfStylableElement element ^String style-name]
  (.setStyleName element style-name))

(defn child-elements
  "The child elements of a node that are instances of the class."
  [^Node node ^Class cls]
  (let [children (.getChildNodes node)]
    (for [i (range (.getLength children))
          :let [child (.item children (int i))]
          :when (instance? cls child)]
      child)))

(defn cell-value
  "The value of a spreadsheet cell: a Double for numeric cells, a String for
  anything else, or nil when blank."
  [^OdfTableCell cell]
  (let [value-type (.getValueType cell)]
    (if (contains? #{"float" "percentage" "currency"} value-type)
      (.getDoubleValue cell)
      (let [s (.getStringValue cell)]
        (when-not (or (nil? s) (= "" s)) s)))))

(defn pt->mm
  "Converts points to whole millimetres."
  ^long [points]
  (Math/round (* 0.3528 (double points))))
