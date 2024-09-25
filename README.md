# Rimborsi scherma

This script helps me automate some of the most tedious work as a computer technician for fencing competitions: making payment documents for all the people involved.

For privacy reasons some files are missing (e.g. `data/JSON/gsa.dt`, `data/templates/template_convocazione.xlsx` and `data/templates/template_rimborso.pdf`).

The program will not work without them, you can create two blank templates (an empty xlsx and a pdf with fillable entries for which the keys are in `data/keys`)

## Schemas

The schema for `data/JSON/gsa.dt` is as follows:

```jsonc
{
    "Arbitri": [
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
    ],
    "Città_Origine": [
        "List",
        "Of",
        "Strings"
    ]
}
```

Values for `LuogoNascita` and `Domicilio` are sometimes null as they are not mandatory to comunicate

The files `*save_conf` and `*.fis_repo` are json like files with a specific schema

The schema for `*save_conf` files is:

```jsonc
{
    "NomeGara": "NameOfCompetition",
    "TipoGara": "TypeOfCompetition", //REG, INTERREG, NAZ
    "CittàGara": "PlaceOfCompetition",
    "IndirizzoGara": "AddressOfCompetition",
    "Tratte": {
        "CityOfOrigin": 999
    },
    "DataGara": "DD-MM-YYYY",
    "DataConv": "DD-MM-YYYY",
    "DataFirma": "DD-MM-YYYY",
    "CostoBenzina": "1.95",
    "ArbitriConv": {
        "xxxxxx": { //FIS Id of summoned Referees
            "Giorni": 2,
            "Extra": false
        } 
    }
}
```

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

Disclaimer: data scraping such as the one used for maps distances is against the TOS, I therefore remove myself from any responsibility of misuse

## Requirements

- [Python 3.x](https://www.python.org/downloads/)
- pip requirements (`pip install -U -r requirements.txt`)
- Internet connection (Google Maps Scraper)
- Excel (Export to pdf, as of now windows only)
- xdg-open (Linux only)
