# VBA_Excel
scripts for Microsoft Excel

##fuzzySearch.bas
This script addresses the problem of searches within cells.  A standard search can highlight a cell but not an individual string of characters or a word.  By crawling along the contents of each cell and comparing chunks of text to chunks of a defined list of words (this was designed to find names in a spreadsheet containing 70K rows including text related to specific meetings). 

fuzzySearch looks inside each cell in a selected range and compares the contents in three letter chunks to items in a master list range.  If a three letter chunk is identified, the chunk shifts one character to the right and compares it.  Depending on the number of matches of three letter chunks, a word will be either unhighlighted (no matches), highlighted yellow (30% match), blue (75% match) or red (100% match).  
