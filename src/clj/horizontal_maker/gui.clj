(ns horizontal-maker.gui
  "The Swing user interface: a window to generate the horizontal and a second
  one to generate the vertical."
  (:require [clojure.java.io :as io]
            [clojure.stacktrace]
            [horizontal-maker.constants :as con]
            [horizontal-maker.formats :as formats]
            [horizontal-maker.sheet :as sheet])
  (:import [java.awt Component Dimension FlowLayout Window]
           [java.awt.event ActionListener WindowAdapter]
           [java.io File]
           [javax.swing AbstractButton BoxLayout JButton JFileChooser JFrame JLabel
            JOptionPane JPanel JScrollPane JSpinner JTextArea JTextField
            SpinnerNumberModel SwingUtilities UIManager WindowConstants]
           [javax.swing.filechooser FileFilter]))

(set! *warn-on-reflection* true)

(defn generate [source width]
  (formats/make-hzd! source width))

;;; Swing helpers

(defn- on-action
  "Calls f (with no arguments) whenever the button is pressed."
  [^AbstractButton button f]
  (.addActionListener button (reify ActionListener
                               (actionPerformed [_ _] (f)))))

(defn- button ^JButton [text f]
  (doto (JButton. ^String text)
    (on-action f)))

(defn- row
  "A panel laying its components out left to right."
  ^JPanel [& components]
  (let [panel (JPanel. (FlowLayout. FlowLayout/LEFT))]
    (doseq [^Component c components]
      (.add panel c))
    panel))

(defn- column
  "A panel laying its components out top to bottom."
  ^JPanel [& components]
  (let [panel (JPanel.)]
    (.setLayout panel (BoxLayout. panel BoxLayout/Y_AXIS))
    (doseq [^Component c components]
      (.add panel c))
    panel))

(defn- frame
  "A window with the content, packed but not yet shown."
  ^JFrame [title content [min-width min-height]]
  (doto (JFrame. ^String title)
    (.setDefaultCloseOperation WindowConstants/DISPOSE_ON_CLOSE)
    (.setMinimumSize (Dimension. min-width min-height))
    (.setContentPane content)
    (.pack)))

(defn- file-filter
  "A file chooser filter accepting directories and files with one of the
  extensions."
  ^FileFilter [description extensions]
  (proxy [FileFilter] []
    (accept [^File f]
      (boolean (or (.isDirectory f)
                   (contains? extensions (sheet/extension f)))))
    (getDescription [] description)))

(defn- stack-trace [^Throwable e]
  (let [writer (java.io.StringWriter.)]
    (.printStackTrace e (java.io.PrintWriter. writer))
    (str writer)))

(defn alert [parent message]
  (JOptionPane/showMessageDialog ^Component parent message))

(defn confirm?
  "Asks a yes/no question."
  [parent message]
  (= JOptionPane/YES_OPTION
     (JOptionPane/showConfirmDialog ^Component parent message "Confirm"
                                    JOptionPane/YES_NO_OPTION
                                    JOptionPane/QUESTION_MESSAGE)))

(defn display-error [parent message e]
  (let [text (doto (JTextArea. ^String (stack-trace e))
               (.setEditable false))
        scroll (doto (JScrollPane. text)
                 (.setPreferredSize (Dimension. 560 280)))]
    (JOptionPane/showMessageDialog ^Component parent
                                   (column (JLabel. ^String message) scroll)
                                   "Error"
                                   JOptionPane/ERROR_MESSAGE)))

;;; File choosers

(defn- choose-input-file
  "Shows an open dialog for the spreadsheet and returns the chosen path, or nil."
  [parent]
  (let [chooser (JFileChooser.)
        spreadsheets (file-filter "Spreadsheets (.xls, .xlsx, .ods)" formats/spreadsheet-extensions)]
    (.addChoosableFileFilter chooser spreadsheets)
    (.setFileFilter chooser spreadsheets)
    (when (= JFileChooser/APPROVE_OPTION (.showOpenDialog chooser ^Component parent))
      (.getPath (.getSelectedFile chooser)))))

(def ^:private document-filters
  {:docx (file-filter "Word Document (.docx)" #{:docx})
   :odt (file-filter "OpenDocument Text (.odt)" #{:odt})})

(defn- document-filter-order
  "The document formats to offer, with the one matching the input first."
  [input-loc]
  (if (= :ods (sheet/extension input-loc))
    [:odt :docx]
    [:docx :odt]))

(defn- with-document-extension
  "The chosen path, given the extension implied by the selected filter unless
  it already names a supported document format."
  [path chosen-format]
  (if (contains? formats/document-extensions (sheet/extension path))
    path
    (str path "." (name chosen-format))))

(defn- choose-document-file
  "Shows a save dialog for the vertical and returns the chosen path, or nil."
  [parent input-loc]
  (let [chooser (JFileChooser.)
        order (document-filter-order input-loc)
        format-for (into {} (map (fn [fmt] [(document-filters fmt) fmt]) order))]
    (.setAcceptAllFileFilterUsed chooser false)
    (doseq [fmt order]
      (.addChoosableFileFilter chooser (document-filters fmt)))
    (.setFileFilter chooser (document-filters (first order)))
    (when (= JFileChooser/APPROVE_OPTION (.showSaveDialog chooser ^Component parent))
      (with-document-extension (.getPath (.getSelectedFile chooser))
                               (get format-for (.getFileFilter chooser) (first order))))))

;;; Windows

(defn show-vertical-dialog [^Window parent input-loc]
  (let [output-field (JTextField. 30)
        book-field (JTextField. 30)
        window (promise)
        browse (button "Browse"
                       #(when-let [path (choose-document-file @window input-loc)]
                          (.setText output-field path)))
        generate (button "Generate"
                         (fn []
                           (try
                             (let [file (io/file (.getText output-field))]
                               (when (or (not (.exists file))
                                         (confirm? @window (str (.getName file) " already exists. Overwrite?")))
                                 (formats/make-vertical! (.getText output-field)
                                                         (formats/read-input input-loc)
                                                         (.getText book-field))
                                 (alert @window "Vertical created.")))
                             (catch Exception e
                               (clojure.stacktrace/print-stack-trace e)
                               (display-error @window "Error generating vertical. Please have someone knowledgeable look at the logs." e)))))
        back (fn []
               (.setVisible ^JFrame @window false)
               (.dispose ^JFrame @window)
               (.setVisible parent true))
        close (button "Back" back)
        ^JFrame f (frame "Make Vertical"
                         (column (row (JLabel. "Save to:") output-field browse)
                                 (row (JLabel. "Book Name:") book-field)
                                 (row generate close))
                         [490 90])]
    (deliver window f)
    (.setDefaultCloseOperation f WindowConstants/DO_NOTHING_ON_CLOSE)
    (.addWindowListener f (proxy [WindowAdapter] []
                            (windowClosing [_] (back))))
    (.setVisible f true)))

(defn- show-main-window []
  (let [input-field (JTextField. 30)
        width-field (JSpinner. (SpinnerNumberModel. (int con/TOTAL-WIDTH) (int 1) Integer/MAX_VALUE (int 100)))
        window (promise)
        browse (button "Browse"
                       #(when-let [path (choose-input-file @window)]
                          (.setText input-field path)))
        done (button "Generate Horizontal"
                     (fn []
                       (try
                         (generate (.getText input-field) (.getValue width-field))
                         (alert @window "Horizontal Created. Please close and re-open your spreadsheet to see the changes.")
                         (catch Exception e
                           (clojure.stacktrace/print-stack-trace e)
                           (display-error @window "Error generating horizontal. Ask someone knowledgeable to see the logs." e)))))
        vertical (button "Generate Vertical"
                         (fn []
                           (.setVisible ^JFrame @window false)
                           (show-vertical-dialog @window (.getText input-field))))
        f (frame "Make Horizontal"
                 (column (row (JLabel. "Input File:") input-field browse)
                         (row (JLabel. "Total Width:") width-field)
                         (row done vertical))
                 [490 90])]
    (deliver window f)
    (.setVisible f true)))

(defn start-gui []
  (try
    (UIManager/setLookAndFeel (UIManager/getSystemLookAndFeelClassName))
    (catch Exception _))
  (SwingUtilities/invokeLater show-main-window))
