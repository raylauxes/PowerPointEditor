# PowerPointEditor
Use python-pptx to edit PowerPoint slides

## Compare run.texts of PowerPoint files:
*python PptComparer.py path/to/file1 path/to/file2*

## Create CSV file that lists run.texts and added a "Revised Run Text" column:
*python PptReviser.py path/to/translated_file path/to/original_file* 

Edit the aforementioned CSV file's "Revised Run Text" Column and run the following command to update the PowerPoint file:
*python PptReplacer.py path/to/translated_file path/to/csv_file*

