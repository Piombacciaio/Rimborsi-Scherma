# Compensi scherma

This script helps me automate some of the most tedious work as a computer technician for fencing competitions: making payment documents for all the people involved.

For privacy reasons some files are missing (e.g. `data/JSON/dt.json` and `data/template.pdf`)

The schema for `data/JSON/dt.json` is a list of dictionaries:

```json
[
    {
        "Nome": "John",
        "Cognome": "Doe",
        "DataNascita": "YYYY-MM-DD",
        "DataRinnovo": "YYYY-MM-DD",
        "Domicilio": "Address",
        "Località": "PlaceOfResidence",
        "LuogoNascita": "PlaceOfBirth",
        "MaschioFemmina": "false",
        "NumFIS": "xxxxxx",
        "Qualifica": "COMPUTERISTA"
    },
]
```

Values for `LuogoNascita` and `Domicilio` are sometimes Null as they are not mandatory to comunicate

`MaschioFemmina` being a bool stored as a string is a leftover from the database export that I will probably remove. (ETA: soon.tm)

The keys for `data/template.pdf` are saved in `data/keys` so they shoud give an idea of what the document is like

Disclaimer : data scraping such as the one used for maps distances is against the TOS, I therefore remove myself from any responsibility

## Requirements

- [Python 3.x](https://www.python.org/downloads/)
- pip requirements (`pip install -U -r requirements.txt`)
- Internet connection (Google Maps Scraper)
