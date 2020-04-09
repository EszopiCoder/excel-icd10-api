# excel-icd10-api
Purpose: Demonstrate ICD-10-CM 2020 API usage within VBA framework


This is an example of how to use the NIH's Clinical Table Search Service API for ICD-10-CM. More information on the API is found [here](https://clinicaltables.nlm.nih.gov/apidoc/icd10cm/v3/doc.html)


The API returns a JSON script which is parsed using code found [here](https://github.com/omegastripes/VBA-JSON-parser). The parsed JSON script is written to the workbook sheet(s).


The current file has two modules:
- `modICD10.bas` contains subs/functions which support searching a single term (code or name)
- `formQueryICD` is a userform which supports searching multiple terms (combination of codes and names)


Compatibility:
- Microsoft Excel 2010+
- Not compatible with Mac OS because this program uses unsupported objects
