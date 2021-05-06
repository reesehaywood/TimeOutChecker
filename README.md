# TimeOutChecker
checks time out documents compared to actual treatment days

There are two versions. The new version reads each date the patient was treated and compares that date to the date in the timeout document. A missing date is added to a list of missing dates for the patient. The list of all patients is then added to the SharePoint list for processing by the therapist later.

# Hints
## Local Copy of Word File
I am copying the file from the server to a local folder before reading. This is to help prevent messing with the file that is in the patients EMR.

## Reading the Tables
Around line 240 in the new version I am reading the tables that are in the Word document. This could be changes to paragraphs. Then a regex can be used to search for specific text in a paragraph.

## Getting the correct documents
You have to edit the getPtDocuments string to search for the document description that you want. Mine is "Timeout". I could be "Consent". The % are SQL for wildcard. Thus, a document like "Treatment Timeout" would be found as well as "Treatment Timeout Hypofraction".
