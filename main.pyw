import datetime as dt, json, math, os, PySimpleGUI as PSG
import data.cf as cf
import urllib.parse
from fillpdf import fillpdfs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from threading import Thread

def icon():
  return b'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'
  
def get_distance(origins: list[str]|str, destination:str, window:PSG.Window, journeys:dict[str,float]) -> None:
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
        window["-OUTPUT-TERMINAL-"].update(f"Exception encountered while fetching distance from {place}: {e}\n", text_color_for_value="yellow", append=True)
        place = urllib.parse.unquote_plus(place)
        journeys[place.upper()] = 0
    window["-OUTPUT-TERMINAL-"].update(f"Loaded Routes\n", text_color_for_value="green", append=True)
    driver.quit()
  except Exception as e:
    window["-OUTPUT-TERMINAL-"].update(f"Error Loading Routes: {e}\n", text_color_for_value="red", append=True)

def load_data(directory:str, window: PSG.Window = None) -> tuple[list[dict], list[str], dict[str, dict], list]:
  #Load data for referees, origins, payments and template
  with open(f"{directory}/data/JSON/dt.json", "r", encoding="utf-8") as f:
    dit = json.load(f)
  with open(f"{directory}/data/JSON/città.json", "r", encoding="utf-8") as f:
    origins = json.load(f)
  with open(f"{directory}/data/JSON/gettoni.json", "r") as f:
    payments = json.load(f)
  form_fields = list(fillpdfs.get_form_fields(f"{directory}/data/template.pdf").keys())
  if window != None: window["-OUTPUT-TERMINAL-"].update(f"Loaded Data\n", text_color_for_value="green", append=True)
  return dit, origins, payments, form_fields

def create_dit_tab(dit:list[dict]) -> list:
  dit_list = ([
    PSG.Checkbox(text="", key=f"-SUMMONED-{person["NumFIS"]}-", s=(1,1)), 
    PSG.Text(f"{person["Cognome"]} {person["Nome"]}",s=(40,1)),
    PSG.Input("", key= f"-DAYS-{person["NumFIS"]}-", s=(5,1)),
    PSG.Checkbox(text="", key=f"-EXTRA-{person["NumFIS"]}-", s=(1,1), pad=(15,0))] for person in dit)
  return [
    [PSG.Text("Conv."), PSG.Push(), PSG.Text("Arbitro"), PSG.Push(), PSG.Text("Giorni "), PSG.Text("Extra"), PSG.Text("   ")],
    [PSG.Column(dit_list, s=(1,300), vertical_scroll_only=True, expand_x=True, scrollable=True, sbar_arrow_color="white", sbar_background_color="grey")],
    [PSG.Text("Pozzo Aggiornato"), 
     PSG.Input("", disabled=True, expand_x=True, key="-UPDATED-REPO-", enable_events=True, disabled_readonly_background_color="gray", disabled_readonly_text_color="white"), 
     PSG.FileBrowse("Apri", file_types=(("FIS_REPO files", "*.fis_repo"),("ALL files", "*.*")), button_color="gray")], #note: fis_repo is a normal json with a specific schema, see README.md
     [PSG.Button("Aggiorna Dati", key="-UPDATE-DIT-", button_color="gray", disabled=True, s=(13,1))]
  ]

def create_view(year:int, month:int, day:int, dit:list[dict]) -> list[list[PSG.TabGroup]]:

  year_list = [x for x in range(year - 1, year + 2)][::-1]

  home_tab = [
    [PSG.Text("Nome Gara", s=(11,1)), PSG.Input("Nome in locandina", key="-COMPETITION-NAME-", s=(50,1))], 
    [PSG.Text("Tipo Gara", s=(11,1)), PSG.Combo(["REG", "INTREG", "NAZ"], "Tipo", s=(8,1), key="-COMPETITION-TYPE-", button_background_color="gray", button_arrow_color="white", enable_events=True)],
    [PSG.Text("Città Gara", s=(11,1)), PSG.Input("Città", key="-COMPETITION-PLACE-", s=(50,1))], 
    [PSG.Text("Indirizzo Gara", s=(11,1)), PSG.Input("Via", key="-COMPETITION-ADDRESS-", s=(37,1)), PSG.Button("Load routes", key="-LOAD-ROUTES-", button_color="gray", pad=(10,0))], 
    [PSG.Text("Data Gara", s=(11,1)), 
    PSG.Combo(["%02d" % x for x in range(1, 32)], default_value="Giorno", key="-COMPETITION-DAY-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="Mese", key="-COMPETITION-MONTH-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(year_list, default_value="Anno", key="-COMPETITION-YEAR-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
    [PSG.Text("Convocazione", s=(11,1)), 
    PSG.Combo(["%02d" % x for x in range(1, 32)], default_value="Giorno", key="-CONVOCATION-DAY-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="Mese", key="-CONVOCATION-MONTH-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(year_list, default_value="Anno", key="-CONVOCATION-YEAR-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
    [PSG.Text("Data firma", s=(11,1)), 
    PSG.Combo(["%02d" % x for x in range(1, 32)], default_value="%02d" % day, key="-SIGN-DAY-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="%02d" % month, key="-SIGN-MONTH-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(year_list, default_value=year, key="-SIGN-YEAR-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
    [PSG.Text("Costo Benzina", s=(11,1)), PSG.Input("1.95", key="-GAS-PRICE-", s=(5,1)), PSG.Text("€/L")],
    [PSG.Button("Export", key="-EXPORT-", disabled=True, bind_return_key=True, button_color="gray"), PSG.Push(), PSG.Button("Reload config", key="-RLD-CFG-", button_color="gray")],
    [PSG.Text("Output")],
    [PSG.Multiline(disabled=True, no_scrollbar=True, autoscroll=True, expand_x=True, auto_refresh=True, s=(1, 5), key="-OUTPUT-TERMINAL-")],
    [PSG.Button("Clear output", key="-CLR-OUT-", button_color="gray")],
    #[PSG.Button("Button", key="-DEBUG-")]
  ]
  dit_tab = create_dit_tab(dit)
  new_dit_tab = [
    [PSG.Text("Dati generali"), PSG.Line()],
    [PSG.Text("Nome", s=(15,1)), PSG.Input("Nome Arbitro", key="-NEW-REFEREE-NAME-", s=(50,1), p=(10,0))],
    [PSG.Text("Cognome", s=(15,1)), PSG.Input("Cognome Arbitro", key="-NEW-REFEREE-SURNAME-", s=(50,1), p=(10,0))],
    [PSG.Text("Femmina", s=(15,1)), PSG.Checkbox(text="", key="-NEW-REFEREE-SEX-")],
    [PSG.Text("Luogo Residenza", s=(15,1)), PSG.Input("Comune residenza", key="-NEW-REFEREE-RESIDENCE-", s=(50,1), p=(10,0))],
    [PSG.Text("Dati anagrafici"), PSG.Line()],
    [PSG.Text("Luogo Nascita", s=(15,1)), PSG.Input("Comune nascita", key="-NEW-REFEREE-BIRTH-PLACE-", s=(50,1), p=(10,0))],
    [PSG.Text("Data Nascita", s=(15,1)), 
     PSG.Combo(["%02d" % x for x in range(1, 32)][::-1], "Giorno", key="-NEW-REFEREE-BIRTH-DAY-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo(["%02d" % x for x in range(1, 13)][::-1], "Mese", key="-NEW-REFEREE-BIRTH-MONTH-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo([x for x in range(year - 80, year)][::-1], "Anno", key="-NEW-REFEREE-BIRTH-YEAR-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0))],
    [PSG.Text("Dati Federazione"), PSG.Line()],
    [PSG.Text("Numero FIS", s=(15,1)), PSG.Input("000000", key="-NEW-REFEREE-FIS-ID-", s=(50,1), p=(10,0))],
    [PSG.Text("Data Rinnovo", s=(15,1)), 
     PSG.Combo(["%02d" % x for x in range(1, 32)][::-1], "Giorno", key="-NEW-REFEREE-RENEWAL-DAY-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo(["%02d" % x for x in range(1, 13)][::-1], "Mese", key="-NEW-REFEREE-RENEWAL-MONTH-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo([x for x in range(year - 80, year)][::-1], "Anno", key="-NEW-REFEREE-RENEWAL-YEAR-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0))],
    [PSG.Text("Qualifica", s=(15,1)), PSG.Combo(["ARBITRO ASP.", "ARBITRO NAZ.", "ARBITRO INT.", "TECNICO ARMI", "COMPUTERISTA", "DIRETTORE TORNEO"], "Qualifica", key="-NEW-REFEREE-ROLE-", button_background_color="gray", button_arrow_color="white", s=(20,1), p=(10,0))],
    [PSG.Button("Nuovo Arbitro", key="-ADD-NEW-REFEREE-", button_color="gray")]
    ]
  default_view = [
    [PSG.TabGroup(
        [
          [PSG.Tab("Home", home_tab)],
          [PSG.Tab("Lista Arbitri", dit_tab)],
          [PSG.Tab("Nuovo Arbitro", new_dit_tab)]
        ]
      )
    ]
  ]

  return default_view

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

  default_view = create_view(current_year, current_month, current_day, dit)

  window = PSG.Window(f"Rimborsi Arbitri | by Piombo Andrea", default_view, icon=icon(), finalize=True, keep_on_top=True)
  
  while True:

    events, values = window.read()
    
    if events == PSG.WIN_CLOSED: break
    
    if events == "-CLR-OUT-": window["-OUTPUT-TERMINAL-"].update("")
    if events == "-COMPETITION-TYPE-": window["-EXPORT-"].update(disabled = False)
    #if events == "-DEBUG-": print("Hi!") #Left this here even if not needed 'cause my best friend (she made this) was extremely proud of her work
    if events == "-RLD-CFG-": dit, origins, payments, form_fields = load_data(current_dir, window)
    if events == "-UPDATED-REPO-": window["-UPDATE-DIT-"].update(disabled=False)
    
    if events == "-LOAD-ROUTES-": 
      journeys = {}
      Thread(target=get_distance, args=[origins, values["-COMPETITION-ADDRESS-"], window, journeys]).start() #Moved get_distance to a separate Thread to avoid freezing the window
    
    if events == "-UPDATE-DIT-": 
      with open(values["-UPDATED-REPO-"], "r", encoding="utf-8") as f:
        new_repo = json.load(f)
      for person in dit:
        updated = next(filter(lambda new_data: new_data['NumFIS'] == person["NumFIS"], new_repo["Tesserati"]), None)
        if updated != None:
          person["DataRinnovo"] = updated["DataRinnovo"]
      with open(f"{current_dir}/data/JSON/dt.json", "w", encoding="utf-8") as f:
        json.dump(dit, f, sort_keys=True, indent=4, ensure_ascii=False)
  
    if events == "-ADD-NEW-REFEREE-":
      new_referee = {}
      new_referee["Cognome"] = values["-NEW-REFEREE-SURNAME-"].upper()
      new_referee["DataNascita"] = f"{values["-NEW-REFEREE-BIRTH-YEAR-"]}-{values["-NEW-REFEREE-BIRTH-MONTH-"]}-{values["-NEW-REFEREE-BIRTH-DAY-"]}"
      new_referee["DataRinnovo"] = f"{values["-NEW-REFEREE-RENEWAL-YEAR-"]}-{values["-NEW-REFEREE-RENEWAL-MONTH-"]}-{values["-NEW-REFEREE-RENEWAL-DAY-"]}"
      new_referee["Località"] = values["-NEW-REFEREE-RESIDENCE-"].upper()
      new_referee["LuogoNascita"] = values["-NEW-REFEREE-BIRTH-PLACE-"].upper() if values["-NEW-REFEREE-BIRTH-PLACE-"] != "" else None
      new_referee["MaschioFemmina"] = values["-NEW-REFEREE-SEX-"]
      new_referee["Nome"] = values["-NEW-REFEREE-NAME-"].upper()
      new_referee["NumFIS"] = values["-NEW-REFEREE-FIS-ID-"]
      new_referee["Qualifica"] = values["-NEW-REFEREE-ROLE-"].upper()
      dit.append(new_referee)
      dit = sorted(dit, key=lambda d: d['Cognome'])
      with open(f"{current_dir}/data/JSON/dt.json", "w", encoding="utf-8") as f:
        json.dump(dit, f, sort_keys=True, indent=4, ensure_ascii=False)
      if new_referee["LuogoNascita"] not in origins and new_referee["LuogoNascita"] != None:
        origins.append(new_referee["LuogoNascita"])
        origins.sort()
        with open(f"{current_dir}/data/JSON/città.json", "w", encoding="utf-8") as f:
          json.dump(origins, f, indent=4, ensure_ascii=False)
      window.close()
      window = PSG.Window(f"Rimborsi Arbitri | by Piombo Andrea", create_view(current_year, current_month, current_day, dit), icon=icon(), finalize=True, keep_on_top=True) #Re-create window to update referees tab 

    if events == "-EXPORT-":
      
      gas = float(str(values["-GAS-PRICE-"]).replace(",", "."))
      competition_name = str(values["-COMPETITION-NAME-"]).upper()
      competition_type = str(values["-COMPETITION-TYPE-"]).upper()
      competition_date = f"{values["-COMPETITION-DAY-"]}/{values["-COMPETITION-MONTH-"]}/{values["-COMPETITION-YEAR-"]}"
      convocation = f"{values["-CONVOCATION-DAY-"]}/{values["-CONVOCATION-MONTH-"]}/{values["-CONVOCATION-YEAR-"]}"
      competition_place = place = str(values["-COMPETITION-PLACE-"]).upper()
      sign_date = f"{values["-SIGN-DAY-"]}/{values["-SIGN-MONTH-"]}/{values["-SIGN-YEAR-"]}"
      for person in dit:
        try:
          if values[f"-SUMMONED-{person["NumFIS"]}-"] == True:
            referee_name = person["Cognome"] + ' ' + person["Nome"]
            window["-OUTPUT-TERMINAL-"].update(f"Adding {referee_name} to export\n", text_color_for_value="green", append=True)
            referee_birthday = person["DataNascita"]
            year, month, day = referee_birthday.split("-")
            referee_birthday = f'{day}/{month}/{year}'
            sex = person["MaschioFemmina"]
            if person["LuogoNascita"] != None:
              referee_birth_place = person["LuogoNascita"] + ', ' + referee_birthday
              referee_tax_code = cf.calcolo_codice(person["Cognome"], person["Nome"], day, month, year, person["LuogoNascita"], sex)
            else:
              referee_birth_place = " "
              referee_tax_code = " "
                  
            try: referee_residence_address = person["Domicilio"] 
            except KeyError: referee_residence_address = " "
                  
            referee_role = person["Qualifica"]
            days = int(values[f"-DAYS-{person["NumFIS"]}-"])
            token_value = payments[competition_type][referee_role]["GETTONE"]
            total_token_value = str(int(days) * int(token_value))
            if values[f"-EXTRA-{person["NumFIS"]}-"] == True: total_token_value = str(int(total_token_value)+token_value)

            travel_distance = math.ceil(journeys[person["Località"].upper()])
            if travel_distance < 50: journey = math.ceil(gas / 10 * 2 * days * travel_distance)
            else: journey = math.ceil(gas / 10 * 2 * travel_distance)
            
            breakfast_number = (1 * days) if travel_distance > 10 else 0
            meal_number = (2 * days) if travel_distance < 100 else (2 * days + 1)
            breakfast_value = payments[competition_type][referee_role]["COLAZIONE"]
            meal_value = payments[competition_type][referee_role]["PRANZO"]
            meals = breakfast_number * breakfast_value + meal_number * meal_value

            nights = days if travel_distance >= 100 else (days -1) if travel_distance >= 50 else 0
            night_value = nights * payments[competition_type][referee_role]["PERNOTTO"]
            total_value = str(float(total_token_value) + float(journey) + float(meals) + float(night_value))
            
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

            fillpdfs.write_fillable_pdf('data/template.pdf',f'Export/{referee_name}.pdf', datadict)
            window["-OUTPUT-TERMINAL-"].update(f"Added {referee_name}\n", text_color_for_value="green", append=True)

        except Exception as e:
          window["-OUTPUT-TERMINAL-"].update(f"Adding {referee_name} raised the following error ({e})\n", text_color_for_value="red", append=True)

if __name__ == "__main__":
  main()
  quit(0)