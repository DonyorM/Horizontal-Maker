# Horizontal Maker

This program is intended to make it easier to create Horizontal charts as used
by the University of Nations School of Biblical Studies and Bible Core Course.

## Download

Get the latest version of the file from
[releases](https://github.com/DonyorM/Horizontal-Maker/releases)

Horizontal Maker requires at least [Java 11](https://adoptium.net/).

## Usage

Create a new spreadsheet and save it as `.xlsx`, `.xls` (Excel) or `.ods`
(LibreOffice / OpenOffice). Enter the horizontal information in the following
format:

| Chapter | Verse | Paragraph Title | End | Segment   | Section   | Division |
|---------|-------|-----------------|-----|-----------|-----------|----------|
| 1       |     1 | Four words only |     | Segment 1 | Section 1 |          |
|         |     5 | Paragraph 2     |     |           |           |          |
|         |    10 | Paragraph 3     |  15 | Segment 2 |           |          |
| 2       |     1 | Paragraph 4     |     |           |           |          |
|         |    13 | Paragraph 5     |  17 |           |           |          |
| 3       |     2 | Paragraph 6     |     | Segment 3 | Section 2 |          |
|         |     9 | Paragraph 7     |     |           |           |          |
|         |    15 | Paragraph 8     |  25 | Segment 4 |           |          |
| 4       |     1 | Paragraph 9     |     |           |           |          |
|         |    12 | Paragraph 10    |  22 |           |           |          |

- Chapter: A label to help you keep track of which chapter you are in. It is
  not used by the program, so you can leave it blank or write whatever you like.
- Verse: The first verse in a paragraph.
- Paragraph Title: The title of the paragraph.
- End: The last verse of a chapter. Leave blank if a chapter does not end in the
  current paragraph.
- Segment: The segment title. Leave blank if a new segment does not begin with
  the current paragraph.
- Section: The section title. Leave blank if a new section does not begin with
  the current paragraph.
- Division: The Division title. Leave blank if a new division does not begin
  with the current paragraph.

Rows whose Verse column isn't a number (such as the heading row) are ignored.

### Horizontal

Save the spreadsheet. Then run the program, select the file you created, and
press "Generate Horizontal". Afterwards, close and reopen the spreadsheet and
look in the "HZD" sheet for the horizontal itself. The first sheet will not be
changed.

"Total Width" is the width of the whole horizontal in Excel column width units
(1/256 of a character). For `.ods` files it is converted to millimetres.

### Vertical

Press "Generate Vertical", choose where to save the document and enter the
name of the book. The vertical can be saved as a Word document (`.docx`) or an
OpenDocument text file (`.odt`); pick the format in the save dialog or type
the extension yourself.

## Building

The project is built with the [Clojure CLI](https://clojure.org/guides/install_clojure)
and `tools.build`.

```sh
clojure -T:build uber          # builds target/horizontal-maker.jar
java -jar target/horizontal-maker.jar

clojure -T:build javac         # compile the Java sources (needed once before :run or :test)
clojure -M:run                 # start the GUI from source
clojure -M:test                # run the tests (or: clojure -T:build run-tests)
```

## Acknowledgements

Software written by Daniel Manila

Thanks to Stephan Hupf and Eric Mueller for creating the original macro for
generating horizontals.
Thanks to the YWAM Rogaland base for teaching me everything on the BCC.

## License

This software is licensed under the MIT License
