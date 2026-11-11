(ns horizontal-maker.constants)

;; The indexes of each item in a row of the input sheet. See
;; horizontal-maker.sheet/read-rows for the row contract.
(def CHAP-IDX "Chapter label. Free-form, only for the user's reference." 0)
(def START-IDX "First verse of the paragraph." 1)
(def PARA-IDX "Paragraph title." 2)
(def END-IDX "Last verse of the chapter, if the chapter ends in this paragraph." 3)
(def SEG-IDX "Segment title, if a new segment starts here." 4)
(def SECT-IDX "Section title, if a new section starts here." 5)
(def DIV-IDX "Division title, if a new division starts here." 6)
(def NUM-COLS "Number of columns in the input sheet." 7)

(def TOTAL-WIDTH
  "The default total width of the horizontal, in Excel column width units
  (1/256 of a character width)."
  27000)

(def MM-PER-CHAR
  "Millimetres per Excel character width, used to convert column widths for
  OpenDocument spreadsheets."
  2.0)

(def DEFAULT-COL-WIDTH-CHARS "Default column width, in characters." 6)

(def PARA-ROW-HEIGHT-PT "Height of the paragraph (or segment) title row, in points." 120)
(def REF-ROW-HEIGHT-PT "Height of the starting-reference row, in points." 27)
