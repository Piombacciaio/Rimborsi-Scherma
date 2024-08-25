# Compensi scherma

This script helps me automate some of the most tedious work as a computer technician for fencing competitions: making payment documents for all the people involved.

For privacy reasons some files are missing (e.g. `data/JSON/dt.json` and `data/template.pdf`)

The schema for `data/JSON/dt.json` is a list of dictionaries:

```jsonc
[
    {
        "Nome": "John", 
        "Cognome": "Doe",
        "DataNascita": "YYYY-MM-DD",
        "DataRinnovo": "YYYY-MM-DD",
        "Domicilio": "Address", //or null
        "Località": "CityOfResidence",
        "LuogoNascita": "ProvinceOfBirth", //or null, used to calculate tax code
        "MaschioFemmina": false, //false if male, true if female
        "NumFIS": "xxxxxx", //6 digits identifier used by Italian Fencing Federation
        "Qualifica": "COMPUTERISTA" //Role of the person. Valid roles are: DIRETTORE TORNEO, COMPUTERISTA, TECNICO ARMI, ARBITRO INT., ARBITRO NAZ., ARBITRO ASP.
    },
]
```

Values for `LuogoNascita` and `Domicilio` are sometimes null as they are not mandatory to comunicate

The keys for `data/template.pdf` are saved in `data/keys` so they shoud give an idea of what the document is like

Disclaimer : data scraping such as the one used for maps distances is against the TOS, I therefore remove myself from any responsibility

## Requirements

- [Python 3.x](https://www.python.org/downloads/)
- pip requirements (`pip install -U -r requirements.txt`)
- Internet connection (Google Maps Scraper)
