import datetime as dt, json, math, os, PySimpleGUI as PSG
import data.cf as cf
import urllib.parse
from fillpdf import fillpdfs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from threading import Thread

def get_distance(origini: list[str]|str, destination:str, window, viaggi):
  try:
    #URL Encode
    url_origini:list[str] = []
    if type(origini) == str:
      url_origini.append(urllib.parse.quote_plus(origini))
    elif type(origini) == list:
      for luogo in origini: url_origini.append(urllib.parse.quote_plus(luogo))
    else:
      raise TypeError(f"Invalid Type for get_distance. Use list or str not {type(origini)}")
    
    url_destinazione:str = urllib.parse.quote_plus(destination)
    
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

    for luogo in url_origini:
      try:
        driver.get(f"https://www.google.com/maps/dir/{luogo}/{url_destinazione}")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@aria-label='Auto']"))) #Wait for page to finish loading
        driver.find_element(By.XPATH, "//img[@aria-label='Auto']").click()
        wait.until(EC.text_to_be_present_in_element((By.CLASS_NAME, "XdKEzd"), "")) #Wait for page to finish loading
        time, route = driver.find_element(By.CLASS_NAME, "XdKEzd").text.split("\n") #Get best routes
        distance, unit = route.split(" ") #time and unit variables are not used in this project
        luogo = urllib.parse.unquote_plus(luogo)
        viaggi[luogo.upper()] = float(distance.replace(",", "."))
      except Exception as e:
        window["-OUTPUT-TERMINAL-"].update(f"Exception encountered while fetching distance from {luogo}: {e}\n", text_color_for_value="yellow", append=True)
        luogo = urllib.parse.unquote_plus(luogo)
        viaggi[luogo.upper()] = 0
    window["-OUTPUT-TERMINAL-"].update(f"Loaded Routes\n", text_color_for_value="green", append=True)
    driver.quit()
  except Exception as e:
    window["-OUTPUT-TERMINAL-"].update(f"Error Loading Routes: {e}\n", text_color_for_value="red", append=True)

def load_data(directory):
  #Caricamento dati su direttori, partenze, compensi e template del modulo
  with open(f"{directory}/data/JSON/dt.json", "r", encoding="utf-8") as f:
    dit = json.load(f)
  with open(f"{directory}/data/JSON/Città.json", "r", encoding="utf-8") as f:
    partenze = json.load(f)
  with open(f"{directory}/data/JSON/gettoni.json", "r") as f:
    rimborsi = json.load(f)
  form_fields = list(fillpdfs.get_form_fields(f"{directory}/data/template.pdf").keys())
  return dit, partenze, rimborsi, form_fields

def create_dit_tab(dit):
  dit_list = ([
    PSG.Checkbox(text="", key=f"-CONVOCATO-{person["NumFIS"]}-", s=(1,1)), 
    PSG.Text(f"{person["Cognome"]} {person["Nome"]}",s=(40,1)),
    PSG.Input("", key= f"-GIORNI-{person["NumFIS"]}-", s=(5,1)),
    PSG.Checkbox(text="", key=f"-EXTRA-{person["NumFIS"]}-", s=(1,1), pad=(15,0))] for person in dit)
  return [
    [PSG.Text("Conv."), PSG.Push(), PSG.Text("Arbitro"), PSG.Push(), PSG.Text("Giorni "), PSG.Text("Extra"), PSG.Text("   ")],
    [PSG.Column(dit_list, s=(1,350), vertical_scroll_only=True, expand_x=True, scrollable=True, sbar_arrow_color="white", sbar_background_color="grey")]
  ]

def create_view(year, month, day, dit):

  year_list = [x for x in range(year - 1, year + 2)][::-1]

  home_tab = [
    [PSG.Text("Nome Gara", s=(11,1)), PSG.Input("Nome in locandina", key="-NOME-GARA-", s=(50,1))], 
    [PSG.Text("Tipo Gara", s=(11,1)), PSG.Combo(["REG", "INTREG", "NAZ"], "Tipo", s=(8,1), key="-TIPO-GARA-", button_background_color="gray", button_arrow_color="white", enable_events=True)],
    [PSG.Text("Città Gara", s=(11,1)), PSG.Input("Città", key="-LUOGO-GARA-", s=(50,1))], 
    [PSG.Text("Indirizzo Gara", s=(11,1)), PSG.Input("Via", key="-INDIRIZZO-GARA-", s=(37,1)), PSG.Button("Load routes", key="-LOAD-ROUTES-", button_color="gray", pad=(10,0))], 
    [PSG.Text("Data Gara", s=(11,1)), 
    PSG.Combo(["%02d" % x for x in range(1, 32)], default_value="Giorno", key="-GIORNO-GARA-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="Mese", key="-MESE-GARA-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(year_list, default_value="Anno", key="-ANNO-GARA-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
    [PSG.Text("Convocazione", s=(11,1)), 
    PSG.Combo(["%02d" % x for x in range(1, 32)], default_value="Giorno", key="-GIORNO-CONV-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="Mese", key="-MESE-CONV-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(year_list, default_value="Anno", key="-ANNO-CONV-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
    [PSG.Text("Data firma", s=(11,1)), 
    PSG.Combo(["%02d" % x for x in range(1, 32)], default_value="%02d" % day, key="-GIORNO-FIRMA-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="%02d" % month, key="-MESE-FIRMA-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
    PSG.Text("/"),
    PSG.Combo(year_list, default_value=year, key="-ANNO-FIRMA-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
    [PSG.Text("Costo Benzina", s=(11,1)), PSG.Input("1.95", key="-COSTO-BENZINA-", s=(5,1)), PSG.Text("€/L")],
    [PSG.Button("Export", key="-EXPORT-", disabled=True, bind_return_key=True, button_color="gray"), PSG.Push(), PSG.Button("Reload config", key="-RLD-CFG-", button_color="gray")],
    [PSG.Text("Output")],
    [PSG.Multiline(disabled=True, no_scrollbar=True, autoscroll=True, expand_x=True, auto_refresh=True, s=(1, 5), key="-OUTPUT-TERMINAL-")],
    [PSG.Button("Clear output", key="-CLR-OUT-", button_color="gray")],
    #[PSG.Button("Button", key="-DEBUG-")]
  ]
  dit_tab = create_dit_tab(dit)
  new_dit_tab = [
    [PSG.Text("Dati generali"), PSG.Line()],
    [PSG.Text("Nome", s=(15,1)), PSG.Input("Nome Arbitro", key="-NOME-NUOVO-DIT-", s=(50,1), p=(10,0))],
    [PSG.Text("Cognome", s=(15,1)), PSG.Input("Cognome Arbitro", key="-COGNOME-NUOVO-DIT-", s=(50,1), p=(10,0))],
    [PSG.Text("Femmina", s=(15,1)), PSG.Checkbox(text="", key="-SESSO-NUOVO-DIT-")],
    [PSG.Text("Luogo Residenza", s=(15,1)), PSG.Input("Comune residenza", key="-LOCALITA-NUOVO-DIT-", s=(50,1), p=(10,0))],
    [PSG.Text("Dati anagrafici"), PSG.Line()],
    [PSG.Text("Luogo Nascita", s=(15,1)), PSG.Input("Comune nascita", key="-NASCITA-NUOVO-DIT-", s=(50,1), p=(10,0))],
    [PSG.Text("Data Nascita", s=(15,1)), 
     PSG.Combo(["%02d" % x for x in range(1, 32)][::-1], "Giorno", key="-GIORNO-NASCITA-NUOVO-DIT-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo(["%02d" % x for x in range(1, 13)][::-1], "Mese", key="-MESE-NASCITA-NUOVO-DIT-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo([x for x in range(year - 80, year)][::-1], "Anno", key="-ANNO-NASCITA-NUOVO-DIT-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0))],
    [PSG.Text("Dati Federazione"), PSG.Line()],
    [PSG.Text("Numero FIS", s=(15,1)), PSG.Input("000000", key="-NUMERO-FIS-NUOVO-DIT-", s=(50,1), p=(10,0))],
    [PSG.Text("Data Rinnovo", s=(15,1)), 
     PSG.Combo(["%02d" % x for x in range(1, 32)][::-1], "Giorno", key="-GIORNO-RINNOVO-NUOVO-DIT-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo(["%02d" % x for x in range(1, 13)][::-1], "Mese", key="-MESE-RINNOVO-NUOVO-DIT-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0)), PSG.Text("/"),
     PSG.Combo([x for x in range(year - 80, year)][::-1], "Anno", key="-ANNO-RINNOVO-NUOVO-DIT-", button_background_color="gray", button_arrow_color="white", s=(6,1), p=(10,0))],
    [PSG.Text("Qualifica", s=(15,1)), PSG.Combo(["ARBITRO ASP.", "ARBITRO NAZ.", "ARBITRO INT.", "TECNICO ARMI", "COMPUTERISTA", "DIRETTORE TORNEO"], "Qualifica", key="-QUALIFICA-NUOVO-DIT-", button_background_color="gray", button_arrow_color="white", s=(20,1), p=(10,0))],
    [PSG.Button("Add referee", key="-NUOVO-DIT-", button_color="gray")]
    ]
  default_view = [
    [PSG.TabGroup(
        [
          [PSG.Tab("Home", home_tab)],
          [PSG.Tab("Nuovo Arbitro", new_dit_tab)],
          [PSG.Tab("Lista Arbitri", dit_tab, key="-LISTA-DIT-")]
        ], key="-PAGES-"
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
  dit, partenze, rimborsi, form_fields = load_data(current_dir)

  default_view = create_view(current_year, current_month, current_day, dit)

  window = PSG.Window(f"Rimborsi Arbitri | by Piombo Andrea", default_view, finalize=True, keep_on_top=True)
  while True:

    events, values = window.read()
    if events == PSG.WIN_CLOSED: break
    if events == "-TIPO-GARA-": window["-EXPORT-"].update(disabled = False)
    if events == "-CLR-OUT-": window["-OUTPUT-TERMINAL-"].update("")
    if events == "-LOAD-ROUTES-": 
      viaggi = {}
      Thread(target=get_distance, args=[partenze, values["-INDIRIZZO-GARA-"], window, viaggi]).start() #Moved get_distance to a separate Thread to avoid freezing the window
    if events == "-RLD-CFG-": dit, partenze, rimborsi, form_fields = load_data(current_dir)
    #if events == "-DEBUG-": print("Hi!") #Left this here even if not needed 'cause my best friend (she made this) was extremely proud of her work
    if events == "-NUOVO-DIT-":
      nuovo_arbitro = {}
      nuovo_arbitro["Cognome"] = values["-COGNOME-NUOVO-DIT-"].upper()
      nuovo_arbitro["DataNascita"] = f"{values["-ANNO-NASCITA-NUOVO-DIT-"]}-{values["-MESE-NASCITA-NUOVO-DIT-"]}-{values["-GIORNO-NASCITA-NUOVO-DIT-"]}"
      nuovo_arbitro["DataRinnovo"] = f"{values["-ANNO-RINNOVO-NUOVO-DIT-"]}-{values["-MESE-RINNOVO-NUOVO-DIT-"]}-{values["-GIORNO-RINNOVO-NUOVO-DIT-"]}"
      nuovo_arbitro["Località"] = values["-LOCALITA-NUOVO-DIT-"].upper()
      nuovo_arbitro["LuogoNascita"] = values["-NASCITA-NUOVO-DIT-"] if values["-NASCITA-NUOVO-DIT-"] != "" else None.upper()
      nuovo_arbitro["MaschioFemmina"] = "false" if not values["-SESSO-NUOVO-DIT-"] else "true"
      nuovo_arbitro["Nome"] = values["-NOME-NUOVO-DIT-"].upper()
      nuovo_arbitro["NumFIS"] = values["-NUMERO-FIS-NUOVO-DIT-"]
      nuovo_arbitro["Qualifica"] = values["-QUALIFICA-NUOVO-DIT-"].upper()
      dit.append(nuovo_arbitro)
      with open(f"{current_dir}/data/JSON/dt.json", "w", encoding="utf-8") as f:
        json.dump(dit, f, sort_keys=True, indent=4, ensure_ascii=False)
      window.close()
      window = PSG.Window(f"Rimborsi Arbitri | by Piombo Andrea", create_view(current_year, current_month, current_day, dit), finalize=True, keep_on_top=True)



    if events == "-EXPORT-":
      
      benzina = float(str(values["-COSTO-BENZINA-"]).replace(",", "."))
      nomegara = str(values["-NOME-GARA-"]).upper()
      tipogara = str(values["-TIPO-GARA-"]).upper()
      datagara = f"{values["-GIORNO-GARA-"]}/{values["-MESE-GARA-"]}/{values["-ANNO-GARA-"]}"
      convocazione = f"{values["-GIORNO-CONV-"]}/{values["-MESE-CONV-"]}/{values["-ANNO-CONV-"]}"
      luogogara = luogo = str(values["-LUOGO-GARA-"]).upper()
      datafirma = f"{values["-GIORNO-FIRMA-"]}/{values["-MESE-FIRMA-"]}/{values["-ANNO-FIRMA-"]}"
      for person in dit:
        try:
          if values[f"-CONVOCATO-{person["NumFIS"]}-"] == True:
            cognomenome = person["Cognome"] + ' ' + person["Nome"]
            window["-OUTPUT-TERMINAL-"].update(f"Adding {cognomenome} to export\n", text_color_for_value="green", append=True)
            datanascita = person["DataNascita"]
            a,m,g = datanascita.split("-")
            datanascita = f'{g}/{m}/{a}'
            mf = True if person["MaschioFemmina"] == "true" else False
            if person["LuogoNascita"] != None:
              nascita = person["LuogoNascita"] + ', ' + datanascita
              codf = cf.calcolo_codice(person["Cognome"], person["Nome"], g, m, a, person["LuogoNascita"], mf)
            else:
              nascita = " "
              codf = " "
                  
            try: domicilio = person["Domicilio"] 
            except KeyError: domicilio = " "
                  
            qualifica = person["Qualifica"]
            giorni = int(values[f"-GIORNI-{person["NumFIS"]}-"])
            gettone = rimborsi[tipogara][qualifica]["GETTONE"]
            gettonetot = str(int(giorni) * int(gettone))
            if values[f"-EXTRA-{person["NumFIS"]}-"] == True: gettonetot = str(int(gettonetot)+gettone)

            km = math.ceil(viaggi[person["Località"].upper()])
            if km < 50: viaggio = math.ceil(benzina / 10 * 2 * giorni * km)
            else: viaggio = math.ceil(benzina / 10 * 2 * km)
            
            numcol = (1 * giorni) if km > 10 else 0
            numpra = (2 * giorni) if km < 100 else (2 * giorni + 1)
            colazione = rimborsi[tipogara][qualifica]["COLAZIONE"]
            pranzo = rimborsi[tipogara][qualifica]["PRANZO"]
            pasti = numcol * colazione + numpra * pranzo

            notti = giorni if km >= 100 else (giorni -1) if km >= 50 else 0
            pernotto = notti * rimborsi[tipogara][qualifica]["PERNOTTO"]
            totale = str(float(gettonetot) + float(viaggio) + float(pasti) + float(pernotto))
            
            numfis = person["NumFIS"]
            datarinnovo = person["DataRinnovo"]
            a,m,g = datarinnovo.split("-")
            rinnovo = f'{g}/{m}/{a}'
            if values[f"-EXTRA-{person["NumFIS"]}-"] == True: giorni = giorni + 1

            datadict = {
              form_fields[0]: cognomenome,
              form_fields[1]: nascita,
              form_fields[2]: domicilio,
              form_fields[3]: codf,
              form_fields[4]: qualifica,
              form_fields[5]: nomegara,
              form_fields[6]: convocazione,
              form_fields[7]: luogogara,
              form_fields[8]: datagara,
              form_fields[9]: giorni,
              form_fields[10]: gettone,
              form_fields[11]: gettonetot,
              form_fields[12]: viaggio,
              form_fields[13]: pasti,
              form_fields[14]: pernotto,
              form_fields[15]: totale,
              form_fields[16]: luogo,
              form_fields[17]: datafirma,
              form_fields[18]: numfis,
              form_fields[19]: rinnovo,
            }

            fillpdfs.write_fillable_pdf('data/template.pdf',f'Export/{cognomenome}.pdf', datadict)
            window["-OUTPUT-TERMINAL-"].update(f"Added {cognomenome}\n", text_color_for_value="green", append=True)

        except Exception as e:
          window["-OUTPUT-TERMINAL-"].update(f"Adding {cognomenome} raised the following error ({e})\n", text_color_for_value="red", append=True)

if __name__ == "__main__":
  main()