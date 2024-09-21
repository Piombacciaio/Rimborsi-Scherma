import calendar, datetime as dt, json, math, os, platform, PySimpleGUI as PSG, subprocess
import data.cf as cf
import urllib.parse
if os.name == "nt": import win32com.client as win32
from fillpdf import fillpdfs
from openpyxl import load_workbook
from openpyxl.styles import Font
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from threading import Thread

def icon():
  "Byte data of the GUI icon"
  return b'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'
  
def get_distance(origins: list[str]|str, destination:str, window:PSG.Window, journeys:dict[str,float]) -> None:
  """Webscraper to get fastest route on google maps"""
  try:
    #URL Encode
    url_origins:list[str] = []
    if type(origins) == str:
      url_origins.append(urllib.parse.quote_plus(origins))
    elif type(origins) == list:
      for place in origins: url_origins.append(urllib.parse.quote_plus(place))
    else:
      raise TypeError(f"Invalid Type for get_distance. Use list or str not {type(origins)}")
    
    url_destinazion:str = urllib.parse.quote_plus(destination)
    
    #Chrome Setup
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-search-engine-choice-screen")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 10)
    
    #Accept cookies
    driver.set_window_position(-2000,0)
    driver.get("https://maps.google.com/")
    driver.find_element(By.CLASS_NAME, "lssxud").click()

    for place in url_origins:
      try:
        driver.get(f"https://www.google.com/maps/dir/{place}/{url_destinazion}")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@aria-label='Auto']"))) #Wait for page to finish loading
        driver.find_element(By.XPATH, "//img[@aria-label='Auto']").click()
        wait.until(EC.text_to_be_present_in_element((By.CLASS_NAME, "XdKEzd"), "")) #Wait for page to finish loading
        time, route = driver.find_element(By.CLASS_NAME, "XdKEzd").text.split("\n") #Get best routes
        distance, unit = route.split(" ") #time and unit variables are not used in this project
        place = urllib.parse.unquote_plus(place)
        journeys[place.upper()] = float(distance.replace(",", "."))
      except Exception as e:
        window["-OUTPUT-TERMINAL-"].update(f"Errore nel calcolo della tratta da {place}: {e}\n", text_color_for_value="yellow", append=True)
        place = urllib.parse.unquote_plus(place)
        journeys[place.upper()] = 0
    window["-OUTPUT-TERMINAL-"].update(f"Tratte caricate correttamente\n", text_color_for_value="green", append=True)
    driver.quit()
  except Exception as e:
    window["-OUTPUT-TERMINAL-"].update(f"Il caricamento delle tratte ha causato il seguente errore {e}\n", text_color_for_value="red", append=True)

def load_data(directory:str) -> tuple[list[dict], list[str], dict[str, dict], list]:
  """Load data for referees, origins, payments and pdf template"""
  with open(f"{directory}/data/json/gsa.dt", "r", encoding="utf-8") as f:
    gsa = json.load(f)
    dit = gsa["Arbitri"]
    origins = gsa["Città_Origine"]
  with open(f"{directory}/data/json/gettoni.json", "r") as f:
    payments = json.load(f)
  form_fields = list(fillpdfs.get_form_fields(f"{directory}/data/templates/template_rimborso.pdf").keys())
  return dit, origins, payments, form_fields

def create_view(year:int, month:int, day:int, dit:list[dict]) -> list[list[PSG.TabGroup]]:

  def empty_line():
    return PSG.Text("void", text_color="black")

  comp_tab = [
    [PSG.Text("Nome Gara", s=(11,1)), PSG.Input("Nome in locandina", key="-COMPETITION-NAME-", s=(55,1))], 
    [PSG.Text("Tipo Gara", s=(11,1)), PSG.Combo(["REG", "INTREG", "NAZ"], "Tipo", s=(8,1), key="-COMPETITION-TYPE-", button_background_color="gray", button_arrow_color="white", enable_events=True)],
    [PSG.Text("Città Gara", s=(11,1)), PSG.Input("Città", key="-COMPETITION-PLACE-", s=(55,1))], 
    [PSG.Text("Indirizzo Gara", s=(11,1)), PSG.Input("Via", key="-COMPETITION-ADDRESS-", s=(41,1)), PSG.Button("Calcola Tratte", key="-LOAD-ROUTES-", button_color="gray", pad=(6,0))], 
    [PSG.Text("Data Gara", s=(11,1)), 
    PSG.Combo(["%02d.%02d" % (x, x + 1) for x in range(1, 31)], default_value="%02d.%02d" % (day, day + 1), key="-COMPETITION-DAY-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="%02d" % month, key="-COMPETITION-MONTH-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo([x for x in range(year - 1, year + 2)][::-1], default_value=year, key="-COMPETITION-YEAR-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
    [PSG.Text("Convocazione", s=(11,1)), 
    PSG.Combo(["%02d" % x for x in range(1, 32)], default_value= "%02d" % (day - 8 if day - 8 > 0 else (calendar.monthrange(year, (month - 1 if month - 1 > 0 else 12))[1] + day - 8)), key="-CONVOCATION-DAY-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="%02d" % (month if day - 8 > 0 else (month - 1 if month - 1 > 0 else 12)), key="-CONVOCATION-MONTH-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo([x for x in range(year - 1, year + 2)][::-1], default_value= (year if day - 8 > 0 and month - 1 > 0 else year - 1) , key="-CONVOCATION-YEAR-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
    [PSG.Text("Data firma", s=(11,1)), 
    PSG.Combo(["%02d" % x for x in range(1, 32)], default_value="%02d" % day, key="-SIGN-DAY-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="%02d" % month, key="-SIGN-MONTH-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo([x for x in range(year - 1, year + 2)][::-1], default_value=year, key="-SIGN-YEAR-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
    [PSG.Text("Costo Benzina", s=(11,1)), PSG.Input("1.95", key="-GAS-PRICE-", s=(5,1)), PSG.Text("€/L")],
    [PSG.Button("Genera", key="-EXPORT-", disabled=True, bind_return_key=True, button_color="gray"), PSG.Button("Vedi Export", key="-VIEW-EXPORT", button_color="gray"), PSG.Push(), PSG.Button("Ricarica Config", key="-RLD-CFG-", button_color="gray")],
    [PSG.VPush()],
    [PSG.Text("Output"), PSG.Line()],
    [PSG.Multiline(disabled=True, autoscroll=True, expand_x=True, auto_refresh=True, s=(1, 6), key="-OUTPUT-TERMINAL-", sbar_arrow_color="white", sbar_background_color="grey")],
    [PSG.Button("Cancella Output", key="-CLR-OUT-", button_color="gray")],
    #[PSG.Button("Button", key="-DEBUG-")]
  ]
  
  tooltip_dit = "Pozzo aggiornato con le nuove date di rinnovo del tesseramento"
  dit_list = ([
    PSG.Checkbox(text="", key=f"-SUMMONED-{person["NumFIS"]}-", s=(1,1)), 
    PSG.Text(f"{person["Cognome"]} {person["Nome"]}",s=(40,1), enable_events=True, key=f"-NAME-{person["NumFIS"]}-"),
    PSG.Input("", key= f"-DAYS-{person["NumFIS"]}-", s=(5,1)),
    PSG.Checkbox(text="", key=f"-EXTRA-{person["NumFIS"]}-", s=(1,1), pad=(17,0))] for person in dit if person["NumFIS"] != "000000")
  dit_tab = [
    [PSG.Text("Conv."), PSG.Push(), PSG.Text("Arbitro"), PSG.Push(), PSG.Text("Giorni "), PSG.Text("Extra"), PSG.Text("   ")],
    [PSG.Column(dit_list, s=(1,305), vertical_scroll_only=True, expand_x=True, scrollable=True, sbar_arrow_color="white", sbar_background_color="grey")],
    [PSG.Button("Conv. Tutti", key="-SUMMON-ALL-", button_color="gray"), PSG.Button("Conv. Nessuno", key="-SUMMON-NONE-", button_color="gray"), 
     PSG.Button("Ins. Giorni Massivo", key="-SUMMON-DAYS-ALL-", tooltip="Solo selezionati", button_color="gray"), PSG.Button("Canc. Giorni Massivo", key="-SUMMON-DAYS-NONE-", tooltip="Solo selezionati", button_color="gray")],
    [PSG.VPush()],
    [PSG.Text("Agg. Rinnovo FIS", tooltip=tooltip_dit), 
     PSG.Input("", disabled=True, key="-UPDATED-REPO-", tooltip=tooltip_dit, enable_events=True, s=(30,1), disabled_readonly_background_color="black", disabled_readonly_text_color="white"), 
     PSG.FileBrowse("Apri", file_types=(("FIS_REPO files", "*.fis_repo"),), tooltip=tooltip_dit, button_color="gray"), PSG.Button("Aggiorna Dati", key="-UPDATE-DIT-", button_color="gray", disabled=True, s=(13,1))],
  ]
  
  new_dit_tab = [
    [PSG.Text("Dati generali"), PSG.Line()],
    [PSG.Text("Nome", s=(15,1)), PSG.Input("", key="-NEW-REFEREE-NAME-", enable_events=True, s=(50,1), p=(10,0))],
    [PSG.Text("Cognome", s=(15,1)), PSG.Input("", key="-NEW-REFEREE-SURNAME-", enable_events=True, s=(50,1), p=(10,0))],
    [PSG.Text("Femmina", s=(15,1)), PSG.Checkbox(text="", key="-NEW-REFEREE-SEX-")],
    [PSG.Text("Luogo Residenza", s=(15,1)), PSG.Input("", key="-NEW-REFEREE-RESIDENCE-", enable_events=True, s=(50,1), p=(10,0))],
    [PSG.Text("Indirizzo*", s=(15,1), tooltip="Valore Opzionale"), PSG.Input("", key="-NEW-REFEREE-ADDRESS-", enable_events=True, s=(50,1), p=(10,0), tooltip="Valore Opzionale")],
    [PSG.Text("*Valore facoltativo, a discrezione di computerista e membro GSA interessato")],
    [PSG.Text("Dati anagrafici"), PSG.Line()],
    [PSG.Text("Luogo Nascita", s=(15,1)), PSG.Input("", key="-NEW-REFEREE-BIRTH-PLACE-", s=(50,1), p=(10,0))],
    [PSG.Text("Data Nascita", s=(15,1)), 
     PSG.Combo(["%02d" % x for x in range(1, 32)], "Giorno", key="-NEW-REFEREE-BIRTH-DAY-", enable_events=True, button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo(["%02d" % x for x in range(1, 13)], "Mese", key="-NEW-REFEREE-BIRTH-MONTH-", enable_events=True, button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo([x for x in range(year - 80, year)][::-1], "Anno", key="-NEW-REFEREE-BIRTH-YEAR-", enable_events=True, button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0))],
    [PSG.Text("Dati Federazione"), PSG.Line()],
    [PSG.Text("Numero FIS", s=(15,1)), PSG.Input("", key="-NEW-REFEREE-FIS-ID-", enable_events=True, s=(50,1), p=(10,0))],
    [PSG.Text("Data Rinnovo", s=(15,1)), 
     PSG.Combo(["%02d" % x for x in range(1, 32)], "Giorno", key="-NEW-REFEREE-RENEWAL-DAY-", enable_events=True, button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo(["%02d" % x for x in range(1, 13)], "Mese", key="-NEW-REFEREE-RENEWAL-MONTH-", enable_events=True, button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo([x for x in range(year - 1, year + 2)][::-1], "Anno", key="-NEW-REFEREE-RENEWAL-YEAR-", enable_events=True, button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0))],
    [PSG.Text("Qualifica", s=(15,1)), PSG.Combo(["ARBITRO ASP.", "ARBITRO NAZ.", "ARBITRO INT.", "TECNICO ARMI", "COMPUTERISTA", "DIRETTORE TORNEO"], "", key="-NEW-REFEREE-ROLE-", enable_events=True, button_background_color="gray", button_arrow_color="white", s=(20,1), p=(10,0))],
    [PSG.VPush()],
    [PSG.Button("Nuovo Arbitro", key="-ADD-NEW-REFEREE-", button_color="gray", disabled=True)]
    ]
  
  combo_edit_text = list(f"{person["NumFIS"]} - {person["Cognome"].upper()} {person["Nome"].upper()}" for person in dit  if person["NumFIS"] != "000000")
  edit_tab = [
    [PSG.Text("Selezione Arbitro"), PSG.Combo(combo_edit_text, "Seleziona", key="-EDIT-REFEREE-CHOICE-", enable_events=True, s=(37,1), button_background_color="gray", button_arrow_color="white"),
     PSG.Button("Elimina Arbitro", key="-EDIT-REFEREE-DEL-", button_color="gray", disabled=True)],
    [PSG.Text("Dati generali"), PSG.Line()],
    [PSG.Text("Nome", s=(15,1)), PSG.Input("", key="-EDIT-REFEREE-NAME-", s=(50,1), p=(10,0), disabled=True, disabled_readonly_background_color="darkgray", disabled_readonly_text_color="white")],
    [PSG.Text("Cognome", s=(15,1)), PSG.Input("", key="-EDIT-REFEREE-SURNAME-", s=(50,1), p=(10,0), disabled=True, disabled_readonly_background_color="darkgray", disabled_readonly_text_color="white")],
    [PSG.Text("Femmina", s=(15,1)), PSG.Checkbox(text="", key="-EDIT-REFEREE-SEX-", disabled=True)],
    [PSG.Text("Luogo Residenza", s=(15,1)), PSG.Input("", key="-EDIT-REFEREE-RESIDENCE-", s=(50,1), p=(10,0), disabled=True, disabled_readonly_background_color="darkgray", disabled_readonly_text_color="white")],
    [PSG.Text("Indirizzo", s=(15,1)), PSG.Input("", key="-EDIT-REFEREE-ADDRESS-", enable_events=True, s=(50,1), p=(10,0), disabled=True, disabled_readonly_background_color="darkgray", disabled_readonly_text_color="white")],
    [PSG.Text("*Valore facoltativo, a discrezione di computerista e membro GSA interessato")],
    [PSG.Text("Dati anagrafici"), PSG.Line()],
    [PSG.Text("Luogo Nascita", s=(15,1)), PSG.Input("", key="-EDIT-REFEREE-BIRTH-PLACE-", s=(50,1), p=(10,0), disabled=True, disabled_readonly_background_color="darkgray", disabled_readonly_text_color="white")],
    [PSG.Text("Data Nascita", s=(15,1)), 
     PSG.Combo(["%02d" % x for x in range(1, 32)], "Giorno", key="-EDIT-REFEREE-BIRTH-DAY-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0), disabled=True), PSG.Text("/"),
     PSG.Combo(["%02d" % x for x in range(1, 13)], "Mese", key="-EDIT-REFEREE-BIRTH-MONTH-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0), disabled=True), PSG.Text("/"),
     PSG.Combo([x for x in range(year - 80, year)][::-1], "Anno", key="-EDIT-REFEREE-BIRTH-YEAR-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0), disabled=True)],
    [PSG.Text("Dati Federazione"), PSG.Line()],
    [PSG.Text("Numero FIS", s=(15,1)), PSG.Text("", key="-EDIT-REFEREE-FIS-ID-", p=(10,0))],
    [PSG.Text("Data Rinnovo", s=(15,1)), 
     PSG.Combo(["%02d" % x for x in range(1, 32)], "Giorno", key="-EDIT-REFEREE-RENEWAL-DAY-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0), disabled=True), PSG.Text("/"),
     PSG.Combo(["%02d" % x for x in range(1, 13)], "Mese", key="-EDIT-REFEREE-RENEWAL-MONTH-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0), disabled=True), PSG.Text("/"),
     PSG.Combo([x for x in range(year - 1, year + 2)][::-1], "Anno", key="-EDIT-REFEREE-RENEWAL-YEAR-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0), disabled=True)],
    [PSG.Text("Qualifica", s=(15,1)), PSG.Combo(["ARBITRO ASP.", "ARBITRO NAZ.", "ARBITRO INT.", "TECNICO ARMI", "COMPUTERISTA", "DIRETTORE TORNEO"], "", key="-EDIT-REFEREE-ROLE-", button_background_color="gray", button_arrow_color="white", s=(20,1), p=(10,0), disabled=True),
     PSG.Push(), PSG.Button("Salva Modifiche", key="-EDIT-REFEREE-SAVE-", button_color="gray", disabled=True)]
  ]
  
  default_view = [
    [PSG.TabGroup(
        [
          [PSG.Tab("Dati Gara", comp_tab)],
          [PSG.Tab("Lista Arbitri", dit_tab)],
          [PSG.Tab("Nuovo Arbitro", new_dit_tab)],
          [PSG.Tab("Modifica Arbitro", edit_tab)],
          [PSG.Tab("Crea fis_repo", [[]])]
        ]
      )
    ],
    [PSG.Button("Salva Configurazione", key="-SAVE-CONFIG-", button_color="gray"), PSG.Button("Carica Configurazione", key="-LOAD-CONFIG-", button_color="gray"),
     PSG.Checkbox("KeepOnTop", default=True, k="-CHANGE-VISIBILITY-", enable_events=True)]
  ]

  return default_view

def fill_summoning_xlsx(window:PSG.Window ,summon_day:str, summon_month:str, summon_year:str, competition_name:str,
                        competition_day:str, competition_month:str, competition_year:str, competition_place:str, 
                        precomp_tech:list, directors:list, comp_techs:list, referees:list, current_dir:str) -> None:

  try:

    window["-OUTPUT-TERMINAL-"].update(f"Creazione di CONVOCAZIONE\n", text_color_for_value="green", append=True)
    if len(referees) > 30:
      window["-OUTPUT-TERMINAL-"].update(f"Troppi arbitri convocati per il template, compilazione manuale richiesta\n", text_color_for_value="yellow", append=True)
      return
    competition_place_cell = competition_place.upper()
    summon_cell = f"{summon_day}-{summon_month}-{summon_year}"
    competition_cell = f"{competition_place_cell} {competition_day}-{competition_month}-{competition_year}"
    precomp_cell = "; ".join(precomp_tech)
    directors_cell = "; ".join(directors)
    comp_techs_cell = "; ".join(comp_techs)

    workbook = load_workbook(f"{current_dir}/data/templates/template_convocazione.xlsx")
    sheet = workbook["template"]

    sheet["F11"] = summon_cell
    sheet["B13"] = competition_name.title()
    sheet["B13"].font = Font(bold=True)
    sheet["B14"] = competition_cell
    sheet["C19"] = precomp_cell
    sheet["C19"].font = Font(bold=True)
    sheet["C20"] = directors_cell
    sheet["C20"].font = Font(bold=True)
    sheet["C21"] = comp_techs_cell
    sheet["C21"].font = Font(bold=True)
    sheet["A37"] = competition_place_cell

    ref_cols = ['B', 'C', 'D']
    ref_start_row = 25
    i = 0

    #Ensure the table is empty
    for col in ref_cols:
      for row in range(ref_start_row, ref_start_row + 10):
        sheet[f"{col}{row}"] = ""

    #Populate table with all ref names
    for col in ref_cols:
      for row in range(ref_start_row, ref_start_row + 10):
        if i < len(referees):
          sheet[f"{col}{row}"] = referees[i]
          i += 1
        else:
          break
    
    if os.name == "nt":
      try:
        workbook.save(f"{current_dir}/data/templates/template_convocazione.xlsx")
        excel = win32.Dispatch('Excel.Application')
        workbook = excel.Workbooks.Open(f"{current_dir}/data/templates/template_convocazione.xlsx")
        sheet = workbook.Worksheets(1)
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = 1
        sheet.PageSetup.Orientation = 2
        sheet.PageSetup.CenterHorizontally = True
        sheet.PageSetup.CenterVertically = True
        pdf_path = f"{current_dir}/Export/CONVOCAZIONE.pdf"
        if os.path.exists(pdf_path): os.remove(pdf_path)
        workbook.ExportAsFixedFormat(0, pdf_path)
        workbook.Close(False)
        window["-OUTPUT-TERMINAL-"].update(f"Compilazione CONVOCAZIONE completata correttamente\n", text_color_for_value="green", append=True)
      finally:
        excel.Quit()
    else:
      workbook.save("Export/CONVOCAZIONE.xlsx")
      window["-OUTPUT-TERMINAL-"].update(f"Compilazione CONVOCAZIONE.xlsx completata correttamente. Procedere all'esportazione in PDF\n", text_color_for_value="yellow", append=True)
  
  except Exception as e:
    window["-OUTPUT-TERMINAL-"].update(f"La creazione di CONVOCAZIONE ha generato il seguente errore: {e}\n", text_color_for_value="red", append=True)

def save_config(window:PSG.Window, current_dir:str, dit:list[dict], values:dict[str], journeys:dict):
  summoned = {}
  for person in dit:
    if person["NumFIS"] != "000000":
      if values[f"-SUMMONED-{person["NumFIS"]}-"] == True:
        summoned[person["NumFIS"]] = {}
        summoned[person["NumFIS"]]["Giorni"] = values[f"-DAYS-{person["NumFIS"]}-"]
        summoned[person["NumFIS"]]["Extra"] = values[f"-EXTRA-{person["NumFIS"]}-"]
  
  try:
    with open(f"{current_dir}/data/json/save_conf", "r", encoding="utf-8") as f:
      savefile = json.load(f)
  except:
    savefile = {}

  savefile["NomeGara"] = values["-COMPETITION-NAME-"].upper()
  savefile["TipoGara"] = values["-COMPETITION-TYPE-"].upper()
  savefile["CittàGara"] = values["-COMPETITION-PLACE-"].upper()
  savefile["IndirizzoGara"] = values["-COMPETITION-ADDRESS-"].title()
  savefile["Tratte"] = journeys
  savefile["DataGara"] = f"{values["-COMPETITION-DAY-"]}-{values["-COMPETITION-MONTH-"]}-{values["-COMPETITION-YEAR-"]}"
  savefile["DataConv"] = f"{values["-CONVOCATION-DAY-"]}-{values["-CONVOCATION-MONTH-"]}-{values["-CONVOCATION-YEAR-"]}"
  savefile["DataFirma"] = f"{values["-SIGN-DAY-"]}-{values["-SIGN-MONTH-"]}-{values["-SIGN-YEAR-"]}"
  savefile["CostoBenzina"] = values["-GAS-PRICE-"]
  savefile["ArbitriConv"] = summoned

  with open(f"{current_dir}/data/json/save_conf", "w", encoding="utf-8") as f:
    json.dump(savefile, f, indent=4, ensure_ascii=False)
      
  window["-OUTPUT-TERMINAL-"].update(f"Configurazione salvata correttamente\n", text_color_for_value="green", append=True)

def main():
  PSG.theme("DarkBlack")

  today = dt.date.today()
  current_year = today.year
  current_month = today.month
  current_day = today.day
  current_dir, _ = str(os.path.realpath(__file__)).replace("\\", "/").rsplit("/", 1)
  
  dit:list[dict]
  origins:list[str]
  payments:dict[dict]
  dit, origins, payments, form_fields = load_data(current_dir)
  journeys = {}

  default_view = create_view(current_year, current_month, current_day, dit)

  window = PSG.Window(f"Rimborsi Arbitri | by Piombo Andrea", default_view, icon=icon(), finalize=True, keep_on_top=True)
  
  while True:

    events:str
    values:dict[str, str|bool]
    events, values = window.read()
    
    if events == PSG.WIN_CLOSED: break
    
    if events == "-CLR-OUT-": window["-OUTPUT-TERMINAL-"].update("")
    if events == "-COMPETITION-TYPE-": window["-EXPORT-"].update(disabled = False)
    #if events == "-DEBUG-": print("Hi!") #Left this here even if not needed 'cause my best friend (she made this) was extremely proud of her work
    if events == "-RLD-CFG-": 
      dit, origins, payments, form_fields = load_data(current_dir)
      window.close()
      window = PSG.Window(f"Rimborsi Arbitri | by Piombo Andrea", create_view(current_year, current_month, current_day, dit), icon=icon(), finalize=True, keep_on_top=True)
      window["-OUTPUT-TERMINAL-"].update(f"Dati aggiornati\n", text_color_for_value="green", append=True)
    if events == "-UPDATED-REPO-": window["-UPDATE-DIT-"].update(disabled=False)
    if events == "-CHANGE-VISIBILITY-": window.keep_on_top_clear() if not values["-CHANGE-VISIBILITY-"] else window.keep_on_top_set()

    if events == "-LOAD-ROUTES-": 
      journeys = {}
      Thread(target=get_distance, args=[origins, values["-COMPETITION-ADDRESS-"], window, journeys]).start() #Moved get_distance to a separate Thread to avoid freezing the window
    
    if events == "-UPDATE-DIT-": 
      with open(values["-UPDATED-REPO-"], "r", encoding="utf-8") as f:
        new_repo = json.load(f)
      count = 0
      for person in dit:
        updated = next(filter(lambda new_data: new_data['NumFIS'] == person["NumFIS"], new_repo["Tesserati"]), None)
        if updated != None:
          person["DataRinnovo"] = updated["DataRinnovo"]
          count += 1
      with open(f"{current_dir}/data/json/gsa.dt", "r", encoding="utf-8") as f:
        gsa = json.load(f)
        gsa["Arbitri"] = dit
      with open(f"{current_dir}/data/json/gsa.dt", "w", encoding="utf-8") as f:
        json.dump(gsa, f, sort_keys=True, indent=4, ensure_ascii=False)
      window["-OUTPUT-TERMINAL-"].update(f"Aggiornamento dei dati di {count} arbitri completato correttamente\n", text_color_for_value="green", append=True)

    if events in ["-NEW-REFEREE-NAME-", "-NEW-REFEREE-SURNAME-", "-NEW-REFEREE-RESIDENCE-", "-NEW-REFEREE-FIS-ID-", "-NEW-REFEREE-RENEWAL-DAY-", "-NEW-REFEREE-RENEWAL-MONTH-", 
                  "-NEW-REFEREE-RENEWAL-YEAR-", "-NEW-REFEREE-ROLE-", "-NEW-REFEREE-BIRTH-DAY-", "-NEW-REFEREE-BIRTH-MONTH-", "-NEW-REFEREE-BIRTH-YEAR-"]:
      disabled_check = not (values["-NEW-REFEREE-NAME-"].strip() != "" and values["-NEW-REFEREE-SURNAME-"].strip() != ""  and 
                            values["-NEW-REFEREE-RESIDENCE-"].strip() != ""  and values["-NEW-REFEREE-FIS-ID-"].strip() != ""  and 
                            values["-NEW-REFEREE-RENEWAL-DAY-"] != "Giorno" and values["-NEW-REFEREE-RENEWAL-MONTH-"] != "Mese" and 
                            values["-NEW-REFEREE-RENEWAL-YEAR-"] != "Anno" and values["-NEW-REFEREE-ROLE-"].strip() != "" and
                            values["-NEW-REFEREE-BIRTH-DAY-"] != "Giorno" and values["-NEW-REFEREE-BIRTH-MONTH-"] != "Mese" and 
                            values["-NEW-REFEREE-BIRTH-YEAR-"] != "Anno" and len(values["-NEW-REFEREE-FIS-ID-"].strip()) == 6)
      window["-ADD-NEW-REFEREE-"].update(disabled=disabled_check)

    if events == "-ADD-NEW-REFEREE-":
      new_referee = {}
      new_referee["Cognome"] = values["-NEW-REFEREE-SURNAME-"].upper().strip()
      new_referee["DataNascita"] = f"{values["-NEW-REFEREE-BIRTH-YEAR-"]}-{values["-NEW-REFEREE-BIRTH-MONTH-"]}-{values["-NEW-REFEREE-BIRTH-DAY-"]}"
      new_referee["DataRinnovo"] = f"{values["-NEW-REFEREE-RENEWAL-YEAR-"]}-{values["-NEW-REFEREE-RENEWAL-MONTH-"]}-{values["-NEW-REFEREE-RENEWAL-DAY-"]}"
      new_referee["Località"] = values["-NEW-REFEREE-RESIDENCE-"].title().strip()
      new_referee["LuogoNascita"] = values["-NEW-REFEREE-BIRTH-PLACE-"].upper().strip() if values["-NEW-REFEREE-BIRTH-PLACE-"] != "" else None
      new_referee["MaschioFemmina"] = values["-NEW-REFEREE-SEX-"]
      new_referee["Nome"] = values["-NEW-REFEREE-NAME-"].upper().strip()
      new_referee["NumFIS"] = values["-NEW-REFEREE-FIS-ID-"].strip()
      new_referee["Qualifica"] = values["-NEW-REFEREE-ROLE-"].upper().strip()
      new_referee["Domicilio"] = values["-NEW-REFEREE-ADDRESS-"].title().strip()
      if new_referee["NumFIS"].isnumeric():
          if not any(ref["NumFIS"] == new_referee["NumFIS"] for ref in dit):
            dit.append(new_referee)
            dit = sorted(dit, key=lambda d: d['Cognome'])
            with open(f"{current_dir}/data/json/gsa.dt", "r", encoding="utf-8") as f:
              gsa = json.load(f)
              gsa["Arbitri"] = dit
            with open(f"{current_dir}/data/json/gsa.dt", "w", encoding="utf-8") as f:
              json.dump(gsa, f, sort_keys=True, indent=4, ensure_ascii=False)
            if new_referee["Località"] not in origins and new_referee["Località"] != "":
              origins.append(new_referee["Località"])
              origins.sort()
              with open(f"{current_dir}/data/json/gsa.dt", "r", encoding="utf-8") as f:
                gsa = json.load(f)
                gsa["Città_Origine"] = origins
              with open(f"{current_dir}/data/json/gsa.dt", "w", encoding="utf-8") as f:
                json.dump(gsa, f, sort_keys=True, indent=4, ensure_ascii=False)
            window.close()
            window = PSG.Window(f"Rimborsi Arbitri | by Piombo Andrea", create_view(current_year, current_month, current_day, dit), icon=icon(), finalize=True, keep_on_top=True)#Re-create window to update referees tab 
            window["-OUTPUT-TERMINAL-"].update(f"Nuovo arbitro aggiunto correttamente\n", text_color_for_value="green", append=True)
          else:
            window["-NEW-REFEREE-FIS-ID-"].update("", background_color="red")
            PSG.PopupQuickMessage("Codice FIS duplicato, correggere e riprovare", background_color="red")
      else:
          window["-NEW-REFEREE-FIS-ID-"].update("", background_color="red")
          PSG.PopupQuickMessage("Formato codice FIS non valido, correggere e riprovare", background_color="red")

    if events == "-EDIT-REFEREE-SAVE-":
      FIS_id, _ = values["-EDIT-REFEREE-CHOICE-"].split(" - ")
      old_ref = next(filter(lambda ref: ref['NumFIS'] == FIS_id, dit), None)
      old_ref["Nome"] = values["-EDIT-REFEREE-NAME-"].upper().strip()
      old_ref["Cognome"] = values["-EDIT-REFEREE-SURNAME-"].upper().strip()
      old_ref["MaschioFemmina"] = values["-EDIT-REFEREE-SEX-"]
      old_ref["Località"] = values["-EDIT-REFEREE-RESIDENCE-"].title().strip()
      old_ref["LuogoNascita"] = values["-EDIT-REFEREE-BIRTH-PLACE-"].upper().strip()
      old_ref["DataNascita"] = f"{values["-EDIT-REFEREE-BIRTH-YEAR-"]}-{values["-EDIT-REFEREE-BIRTH-MONTH-"]}-{values["-EDIT-REFEREE-BIRTH-DAY-"]}"
      old_ref["DataRinnovo"] = f"{values["-EDIT-REFEREE-RENEWAL-YEAR-"]}-{values["-EDIT-REFEREE-RENEWAL-MONTH-"]}-{values["-EDIT-REFEREE-RENEWAL-DAY-"]}"
      old_ref["Qualifica"] = values["-EDIT-REFEREE-ROLE-"].upper().strip()
      old_ref["Domicilio"] = values["-EDIT-REFEREE-ADDRESS-"].title().strip()
      with open(f"{current_dir}/data/json/gsa.dt", "r", encoding="utf-8") as f:
        gsa = json.load(f)
        gsa["Arbitri"] = dit
      with open(f"{current_dir}/data/json/gsa.dt", "w", encoding="utf-8") as f:
        json.dump(gsa, f, sort_keys=True, indent=4, ensure_ascii=False)
      window.close()
      window = PSG.Window(f"Rimborsi Arbitri | by Piombo Andrea", create_view(current_year, current_month, current_day, dit), icon=icon(), finalize=True, keep_on_top=True)
      window["-OUTPUT-TERMINAL-"].update(f"Arbitro modificato correttamente\n", text_color_for_value="green", append=True)

    if events == "-EDIT-REFEREE-DEL-":
      FIS_id, _ = str(values["-EDIT-REFEREE-CHOICE-"]).split(" - ")
      edit_ref_chosen = next(filter(lambda ref: ref['NumFIS'] == FIS_id, dit), None)
      dit.remove(edit_ref_chosen)
      with open(f"{current_dir}/data/json/gsa.dt", "r", encoding="utf-8") as f:
        gsa = json.load(f)
        gsa["Arbitri"] = dit
      with open(f"{current_dir}/data/json/gsa.dt", "w", encoding="utf-8") as f:
        json.dump(gsa, f, sort_keys=True, indent=4, ensure_ascii=False)
      window.close()
      window = PSG.Window(f"Rimborsi Arbitri | by Piombo Andrea", create_view(current_year, current_month, current_day, dit), icon=icon(), finalize=True, keep_on_top=True)
      window["-OUTPUT-TERMINAL-"].update(f"Arbitro eliminato correttamente\n", text_color_for_value="green", append=True)

    if events == "-EDIT-REFEREE-CHOICE-":
      FIS_id, _ = values["-EDIT-REFEREE-CHOICE-"].split(" - ")
      edit_ref_chosen = next(filter(lambda ref: ref['NumFIS'] == FIS_id, dit), None)
      window["-EDIT-REFEREE-DEL-"].update(disabled=False)
      window["-EDIT-REFEREE-NAME-"].update(edit_ref_chosen["Nome"], disabled=False)
      window["-EDIT-REFEREE-SURNAME-"].update(edit_ref_chosen["Cognome"], disabled=False)
      window["-EDIT-REFEREE-SEX-"].update(edit_ref_chosen["MaschioFemmina"], disabled=False)
      window["-EDIT-REFEREE-RESIDENCE-"].update(edit_ref_chosen["Località"], disabled=False)
      window["-EDIT-REFEREE-BIRTH-PLACE-"].update(edit_ref_chosen["LuogoNascita"], disabled=False)
      window["-EDIT-REFEREE-ADDRESS-"].update(edit_ref_chosen["Domicilio"], disabled=False)
      year,month,day = edit_ref_chosen["DataNascita"].split("-")
      window["-EDIT-REFEREE-BIRTH-DAY-"].update(day, disabled=False)
      window["-EDIT-REFEREE-BIRTH-MONTH-"].update(month, disabled=False)
      window["-EDIT-REFEREE-BIRTH-YEAR-"].update(year, disabled=False)
      window["-EDIT-REFEREE-FIS-ID-"].update(edit_ref_chosen["NumFIS"])
      year,month,day = edit_ref_chosen["DataRinnovo"].split("-")
      window["-EDIT-REFEREE-RENEWAL-DAY-"].update(day, disabled=False)
      window["-EDIT-REFEREE-RENEWAL-MONTH-"].update(month, disabled=False)
      window["-EDIT-REFEREE-RENEWAL-YEAR-"].update(year, disabled=False)
      window["-EDIT-REFEREE-ROLE-"].update(edit_ref_chosen["Qualifica"], disabled=False)
      window["-EDIT-REFEREE-SAVE-"].update(disabled=False)

    if events.startswith("-NAME-"):
      summon = events.replace("-NAME-", "-SUMMONED-")
      window[summon].update(not values[summon])

    if events == "-EXPORT-":
      if journeys != {}:
        if len(journeys) == len(origins):
          comp_techs = []
          precomp_tech = []
          directors = []
          referees = []

          gas:float = float(values["-GAS-PRICE-"].replace(",", "."))
          competition_name:str = values["-COMPETITION-NAME-"].upper()
          competition_type:str = values["-COMPETITION-TYPE-"].upper()
          competition_date:str = f"{values["-COMPETITION-DAY-"]}/{values["-COMPETITION-MONTH-"]}/{values["-COMPETITION-YEAR-"]}"
          convocation:str = f"{values["-CONVOCATION-DAY-"]}/{values["-CONVOCATION-MONTH-"]}/{values["-CONVOCATION-YEAR-"]}"
          competition_place:str
          place:str
          competition_place = place = values["-COMPETITION-PLACE-"].upper()
          sign_date:str = f"{values["-SIGN-DAY-"]}/{values["-SIGN-MONTH-"]}/{values["-SIGN-YEAR-"]}"
          for person in dit:
            if person["NumFIS"] != "000000":
              try:
                if values[f"-SUMMONED-{person["NumFIS"]}-"] == True:
                  referee_name:str = person["Cognome"] + ' ' + person["Nome"]
                  window["-OUTPUT-TERMINAL-"].update(f"Creazione di {referee_name}\n", text_color_for_value="green", append=True)
                  referee_birthday:str = person["DataNascita"]
                  year:str
                  month:str
                  day:str
                  year, month, day = referee_birthday.split("-")
                  referee_birthday:str = f'{day}/{month}/{year}'
                  sex:bool = person["MaschioFemmina"]
                  if person["LuogoNascita"] != None:
                    referee_birth_place:str = person["LuogoNascita"] + ', ' + referee_birthday
                    referee_tax_code:str = cf.calcolo_codice(person["Cognome"], person["Nome"], day, month, year, person["LuogoNascita"], sex)
                  else:
                    referee_birth_place:str = " "
                    referee_tax_code:str = " "
                        
                  try: referee_residence_address:str = person["Domicilio"] 
                  except KeyError: referee_residence_address:str = " "
                        
                  referee_role:str = person["Qualifica"]
                  if referee_role in ["ARBITRO ASP.", "ARBITRO NAZ.",  "ARBITRO INT."]: referees.append(referee_name)
                  if referee_role == "DIRETTORE TORNEO": directors.append(referee_name)
                  if referee_role == "COMPUTERISTA": comp_techs.append(referee_name)
                  if values[f"-EXTRA-{person["NumFIS"]}-"] == True and referee_role in ["COMPUTERISTA", "DIRETTORE TORNEO"]: precomp_tech.append(referee_name)
                  days = int(values[f"-DAYS-{person["NumFIS"]}-"])
                  token_value:int = payments[competition_type][referee_role]["GETTONE"]
                  total_token_value:int = int(days * token_value)
                  if values[f"-EXTRA-{person["NumFIS"]}-"] == True: 
                    total_token_value:int = total_token_value + token_value

                  travel_distance:int = math.ceil(journeys[person["Località"].upper()])
                  if travel_distance < 50: 
                    journey = math.ceil(gas / 10 * 2 * days * travel_distance)
                  else: 
                    journey = math.ceil(gas / 10 * 2 * travel_distance)

                  breakfast_number = (1 * days) if travel_distance > 10 else 0
                  meal_number = (2 * days) if travel_distance < 100 else (2 * days + 1)
                  breakfast_value = payments[competition_type][referee_role]["COLAZIONE"]
                  meal_value = payments[competition_type][referee_role]["PRANZO"]
                  meals = int(breakfast_number * breakfast_value + meal_number * meal_value)

                  nights = days if travel_distance >= 100 else (days -1) if travel_distance >= 50 else 0
                  night_value:int = nights * payments[competition_type][referee_role]["PERNOTTO"]
                  total_value = str(total_token_value + journey + meals + night_value)
                  try:
                    total_value, _ = total_value.split(".0")
                  except ValueError:
                    pass
                  
                  FIS_id = person["NumFIS"]
                  datarinnovo = person["DataRinnovo"]
                  year, month, day = datarinnovo.split("-")
                  renewal_date = f'{day}/{month}/{year}'
                  if values[f"-EXTRA-{person["NumFIS"]}-"] == True: days = days + 1

                  datadict = {
                    form_fields[0]: referee_name,
                    form_fields[1]: referee_birth_place,
                    form_fields[2]: referee_residence_address,
                    form_fields[3]: referee_tax_code,
                    form_fields[4]: referee_role,
                    form_fields[5]: competition_name,
                    form_fields[6]: convocation,
                    form_fields[7]: competition_place,
                    form_fields[8]: competition_date,
                    form_fields[9]: days,
                    form_fields[10]: token_value,
                    form_fields[11]: total_token_value,
                    form_fields[12]: journey,
                    form_fields[13]: meals,
                    form_fields[14]: night_value,
                    form_fields[15]: total_value,
                    form_fields[16]: place,
                    form_fields[17]: sign_date,
                    form_fields[18]: FIS_id,
                    form_fields[19]: renewal_date,
                  }

                  fillpdfs.write_fillable_pdf('data/templates/template_rimborso.pdf',f'Export/{referee_name}.pdf', datadict)
                  window["-OUTPUT-TERMINAL-"].update(f"Creazione di {referee_name} completata correttamente\n", text_color_for_value="green", append=True)

              except Exception as e:
                window["-OUTPUT-TERMINAL-"].update(f"La creazione di {referee_name} ha generato il seguente errore: {e}\n", text_color_for_value="red", append=True)

          fill_summoning_xlsx(window, values["-CONVOCATION-DAY-"], values["-CONVOCATION-MONTH-"], values["-CONVOCATION-YEAR-"],
                              values["-COMPETITION-NAME-"], values["-COMPETITION-DAY-"], values["-COMPETITION-MONTH-"], values["-COMPETITION-YEAR-"],
                              values["-COMPETITION-PLACE-"], precomp_tech, directors, comp_techs, referees, current_dir)
          save_config(window, current_dir, dit, values, journeys)
        else:
         window["-OUTPUT-TERMINAL-"].update(f"Impossibile esportare, attendere la fine del calcolo tratte\n", text_color_for_value="yellow", append=True)
      else:
        window["-OUTPUT-TERMINAL-"].update(f"Impossibile esportare, nessuna tratta calcolata\n", text_color_for_value="red", append=True)

    if events == "-SAVE-CONFIG-":
      try:
       journeys = journeys
      except UnboundLocalError:
        journeys = {}
      save_config(window, current_dir, dit, values, journeys)

    if events == "-LOAD-CONFIG-":
      with open(f"{current_dir}/data/json/save_conf", "r", encoding="utf-8") as f:
        savefile = json.load(f)
      
      window["-COMPETITION-NAME-"].update(savefile["NomeGara"])
      window["-COMPETITION-TYPE-"].update(savefile["TipoGara"])
      window["-COMPETITION-PLACE-"].update(savefile["CittàGara"])
      window["-COMPETITION-ADDRESS-"].update(savefile["IndirizzoGara"])
      window["-EXPORT-"].update(disabled=False)
      journeys = savefile["Tratte"]
      
      day,month,year = savefile["DataGara"].split("-")
      window["-COMPETITION-DAY-"].update(day)
      window["-COMPETITION-MONTH-"].update(month)
      window["-COMPETITION-YEAR-"].update(year)
      
      day,month,year = savefile["DataConv"].split("-")
      window["-CONVOCATION-DAY-"].update(day)
      window["-CONVOCATION-MONTH-"].update(month)
      window["-CONVOCATION-YEAR-"].update(year)
      
      day,month,year = savefile["DataFirma"].split("-")
      window["-SIGN-DAY-"].update(day)
      window["-SIGN-MONTH-"].update(month)
      window["-SIGN-YEAR-"].update(year)
      
      window["-GAS-PRICE-"].update(savefile["CostoBenzina"])

      for numfis, vals in savefile["ArbitriConv"].items():
        window[f"-SUMMONED-{numfis}-"].update(value=True)
        window[f"-DAYS-{numfis}-"].update(value=vals["Giorni"])
        window[f"-EXTRA-{numfis}-"].update(value=vals["Extra"])

    if events == "-SUMMON-ALL-":
      for person in dit:
        window[f"-SUMMONED-{person["NumFIS"]}-"].update(True) if person["NumFIS"] != "000000" else ""

    if events == "-SUMMON-NONE-":
      for person in dit:
        window[f"-SUMMONED-{person["NumFIS"]}-"].update(False) if person["NumFIS"] != "000000" else ""
    
    if events == "-SUMMON-DAYS-ALL-":
      mass_days = PSG.popup_get_text("Inserire il numero di giorni della convocazione.\nAssicurarsi di aver selezionato tutti gli arbitri desiderati.", title="Inserimento Giorni Massivo", icon=icon(), button_color="gray", keep_on_top=True)
      if mass_days != None:
        if mass_days.strip() != "" and mass_days.isnumeric():
          mass_days = mass_days.strip()
          for person in dit:
            if person["NumFIS"] != "000000":
              if values[f"-SUMMONED-{person["NumFIS"]}-"]:
                window[f"-DAYS-{person["NumFIS"]}-"].update(value=mass_days)
        else:
          PSG.PopupQuickMessage("Valore giorni non valido, correggere e riprovare", background_color="red")
    
    if events == "-SUMMON-DAYS-NONE-":
      for person in dit:
        if person["NumFIS"] != "000000":
          if values[f"-SUMMONED-{person["NumFIS"]}-"]:
            window[f"-DAYS-{person["NumFIS"]}-"].update(value="")

    if events == "-VIEW-EXPORT":
      system = platform.system()
      path = f"{current_dir}/Export"
      if system == "Windows":
        os.startfile(path)
      elif system == "Darwin":  # macOS
        subprocess.Popen(['open', path])
      elif system == "Linux":
        subprocess.Popen(['xdg-open', os.path.dirname(path)])
      else:
        PSG.PopupQuickMessage(f"Unsupported OS: {system}", background_color="red")

if __name__ == "__main__":
  main()
  quit(0)