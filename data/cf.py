import json, os

BIRTH_MONTH = {
  "01": "A",
  "02": "B",
  "03": "C",
  "04": "D",
  "05": "E",
  "06": "H",
  "07": "L",
  "08": "M",
  "09": "P",
  "10": "R",
  "11": "S",
  "12": "T"}
CIN_CONVERSION = {
  0 : "A",
  1 : "B",
  2 : "C",
  3 : "D",
  4 : "E",
  5 : "F",
  6 : "G",
  7 : "H",
  8 : "I",
  9 : "J",
  10 : "K",
  11 : "L",
  12 : "M",
  13 : "N",
  14 : "O",
  15 : "P",
  16 : "Q",
  17 : "R",
  18 : "S",
  19 : "T",
  20 : "U",
  21 : "V",
  22 : "W",
  23 : "X",
  24 : "Y",
  25 : "Z"}
EVEN_CONVERSION = {
  "0": 0,
  "1": 1,
  "2": 2,
  "3": 3,
  "4": 4,
  "5": 5,
  "6": 6,
  "7": 7,
  "8": 8,
  "9": 9,
  "A": 0,
  "B": 1,
  "C": 2,
  "D": 3,
  "E": 4,
  "F": 5,
  "G": 6,
  "H": 7,
  "I": 8,
  "J": 9,
  "K": 10,
  "L": 11,
  "M": 12,
  "N": 13,
  "O": 14,
  "P": 15,
  "Q": 16,
  "R": 17,
  "S": 18,
  "T": 19,
  "U": 20,
  "V": 21,
  "W": 22,
  "X": 23,
  "Y": 24,
  "Z": 25}
ODD_CONVERSION = {
  "0": 1,
  "1": 0,
  "2": 5,
  "3": 7,
  "4": 9,
  "5": 13,
  "6": 15,
  "7": 17,
  "8": 19,
  "9": 21,
  "A": 1,
  "B": 0,
  "C": 5,
  "D": 7,
  "E": 9,
  "F": 13,
  "G": 15,
  "H": 17,
  "I": 19,
  "J": 21,
  "K": 2,
  "L": 4,
  "M": 18,
  "N": 20,
  "O": 11,
  "P": 3,
  "Q": 6,
  "R": 8,
  "S": 12,
  "T": 14,
  "U": 16,
  "V": 10,
  "W": 22,
  "X": 25,
  "Y": 24,
  "Z": 23}
VOWELS = "AEIOU"
CONSONANTS = "BCDFGHJKLMNPQRSTVWXYZ"
DIGITS = "0123456789"

current_dir, _ = str(os.path.realpath(__file__)).replace("\\", "/").rsplit("/data")

with open(f"{current_dir}/data/JSON/municipalities.json", "r", encoding="utf-8") as codes: 
  CODES = json.load(codes)

def calculate_name_chars(name:str):
  if len(name) <= 3:
    code = ""

    for character in name:
      if character in CONSONANTS:
        code += character

    for character in name:
      if character in VOWELS:
        code += character

    return code + ("X" * (3 - len(name)))
  
  else:
    consonants = ""
    code = ""

    for character in name:
      if character in CONSONANTS and len(code) < 3:
        consonants += character

    if len(consonants) > 3:
      code = consonants[0] + consonants[2] + consonants[3]
    else:
      for character in consonants:
        if len(code) < 3:
          code += character

    for character in name:
      if character in VOWELS and len(code) < 3:
        code += character

    return code
def calculate_surname_chars(name:str):
  if len(name) <= 3:
    code = ""

    for character in name:
      if character in CONSONANTS:
        code += character

    for character in name:
      if character in VOWELS:
        code += character

    return code + ("X" * (3 - len(name)))
  
  else:
    code = ""

    for character in name:
      if character in CONSONANTS and len(code) < 3:
        code += character

    for character in name:
      if character in VOWELS and len(code) < 3:
        code += character

    return code
def calculate_cin_char(partial_code:str):
  even_positions = partial_code[1::2]
  odd_positions = partial_code[0::2]
  
  even_sum = 0
  for character in even_positions:
    value = EVEN_CONVERSION[character]
    even_sum += value

  odd_sum = 0
  for character in odd_positions:
    value = ODD_CONVERSION[character]
    odd_sum += value

  remainder = (even_sum + odd_sum) % 26
  cin_code = CIN_CONVERSION[remainder]
  return cin_code

def calcolo_codice(cognome:str, nome:str, giorno:str, mese:str, anno:str, luogo:str, mf:bool):
  cognome = cognome.upper()
  nome = nome.upper()
  luogo = luogo.lower()

  cod_cognome = calculate_surname_chars(cognome)
  cod_nome = calculate_name_chars(nome)
  municipalità:str = CODES[luogo]["codice_catastale"]
  cod_anno = anno[2:]
  cod_mese = BIRTH_MONTH[mese]
  if mf == True:
    cod_giorno = str(int(giorno) + 40)
  else:
    cod_giorno = giorno

  cod_parziale = cod_cognome+cod_nome+cod_anno+cod_mese+cod_giorno+municipalità
  cin = calculate_cin_char(cod_parziale)
  cod_completo = cod_parziale + cin

  return cod_completo