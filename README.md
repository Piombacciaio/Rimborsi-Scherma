# Rimborsi scherma

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

Disclaimer: data scraping such as the one used for maps distances is against the TOS, I therefore remove myself from any responsibility of misuse

The schema for `*.fis_repo` files is:

```jsonc
{   
    "Tesserati": [
        {
            "Nome": "JHON", //Name of the person
            "Cognome": "DOE", //Surname of the person
            "CatAtleti": "0", //0-9 number used to identify the category
            "CodSoc": "xxxxxx", //2-6 digits number to uniquely identify any club affiliated with the Italian Fencing Federation
            "DataNascita": "YYYY-MM-DD", //Date of birth
            "DataRinnovo": "YYYY-MM-DD", //Date of renewal
            "Località": "CityOfResidence",
            "LuogoNascita": "ProvinceOfBirth", //or null if born out of Italy
            "MaschioFemmina": false, //false if male, true if female
            "NumFIS": "xxxxxx" //6 digits number to uniquely identify any person affiliated with the Italian Fencing Federation
        },
    ],
    "Societa": [
        {
            "CodSoc": "xxxxxx", //2-6 digits number to uniquely identify any club affiliated with the Italian Fencing Federation
            "DataScadenza": "YYYY-MM-DD", //Date of expiry of affiliation
            "Denominazione": "FullNameOfClub",
            "Località": "ProvinceOfOperation",
            "Sigla": "ABCDE" //4-5 characters to uniquely identify any club during competitions
        },
    ]
}
```

## Requirements

- [Python 3.x](https://www.python.org/downloads/)
- pip requirements (`pip install -U -r requirements.txt`)
- Internet connection (Google Maps Scraper)
