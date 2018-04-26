# Horizontal Maker

This program is intended to make it easier to create Horizontal charts as used
by the University of Nations School of Biblical Studies and Bible Core Course.

## Download

Get the latest version of the file from
[releases](https://github.com/DonyorM/Horizontal-Maker/releases)

Horizontal Maker requires at least [Java 6](https://java.com/en/download/).

## Usage

Create a new spreadsheet and save as .xls or .xlsx (other formats may work, but
are not supported.) Enter in the horizontal information in the following format:

| Verse | Paragraph Title | End | Segment   | Section   | Division |
|-------|-----------------|-----|-----------|-----------|----------|
|     1 | Four words only |     | Segment 1 | Section 1 |          |
|     5 | Paragraph 2     |     |           |           |          |
|    10 | Paragraph 3     |  15 | Segment 2 |           |          |
|     1 | Paragraph 4     |     |           |           |          |
|    13 | Paragraph 5     |  17 |           |           |          |
|     2 | Paragraph 6     |     | Segment 3 | Section 2 |          |
|     9 | Paragraph 7     |     |           |           |          |
|    15 | Paragraph 8     |  25 | Segment 4 |           |          |
|     1 | Paragraph 9     |     |           |           |          |
|    12 | Paragraph 10    |  22 |           |           |          |

- Verse: The first verse in a paragraph.
- Paragraph Title: The title of the paragraph.
- End: The last verse of a chapter. Leave blank if a chapter does not end in the
  current paragraph.
- Segment: The segment title. Leave blank if a new segment does not begin with
  the current paragraph.
- Section: The section title. Leave blank if a new section does not begin with
  the current paragraph.
- Division: The Division title. Leave blank if a new section does not begin
  with the current paragraph.
  
Save the excel file. Then run the program and select the excel file you created, and press generate.
Afterwards, close and reopen the excel file and look in the "HZD" sheet for the
horizontal itself. The first sheet will not be changed.
  
## Acknowledgements

Software written by Daniel Manila

Thanks to Stephan Hupf and Eric Mueller for creating the original macro for
generating horizontals.
Thanks to the YWAM Rogaland base for teaching me everything on the BCC.

## License

This software is licensed under the MIT License
