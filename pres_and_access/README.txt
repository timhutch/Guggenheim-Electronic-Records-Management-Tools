pres_and_access.vbs (VBScript)

Requires: Microsoft Office 2007+, Windows XP or greater

Description:
This VBScript uses MS Office to batch convert select formats into formats well-suited for access and preservation.  It uses the file naming structure in use with Archivematica, and thus should work well with Archivematica’s manual file normalization process.  For example, given the following set of files:

\Correspondence\
   \Email1.rtf
   \Email2.wpd
   \Email3.doc
   \Email4.docx
   \Letters
      \form-letter.mlm
\Installation budget
   \Budget1.xls
\Presentations
   \Pres1.ppt
   \old
      \Pres_old.ppt

Will produce the following set of files:

{all files above, plus the following new directory}
\manualNormalization
   \access
      \Correspondence
         Email1.pdf
         Email2.pdf
         Email3.pdf
         Email4.pdf
         \Letters
            form-letter.pdf
      \Installation budget
      \Presentations
         Pres1.pdf
         \old
            Pres_old.pdf
   \preservation
      \Correspondence
         Email2.docx
         Email3.docx
         \Letters
            form-letter.docx
      \Installation budget
          Budget1.xlsx
      \Presentations
         Pres1.pptx
         \old
            Pres_old.pptx


As you can see from the above example, the script uses the following as preservation and access formats:

Preservation:
DOC -> DOCX
PPT -> PPTX
XLS -> XLSX
XLSX -> (no change)
DOCX -> (no change)
PPTX -> (no change)
WPD -> DOCX
MLM -> DOCX
RTF -> (no change)

Access:
DOC -> PDF
PPT -> PDF
XLS -> XLSX
XLSX -> (no change)
DOCX -> PDF
PPTX -> PDF
WPD -> PDF
MLM -> PDF
RTF -> PDF



