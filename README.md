stoic-xlsx
==========

Importer: turn an xlsx into stoic's internal json format:
```json
{
  name: 'spreadsheetName'
  timezone: '',
  ...
  "sheets": {
    "Fields": {
      "values": [
        [
          "A1 ...",
          "B1 ..."
        ],
      },
      "numberFormats": [
        [  ]
      ],
      "notes": [
        [ "text of the comments on A1", "..." ]
      ],
    }
}
```
Exporter: turn an xlsx into 
