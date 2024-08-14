import json, datetime as dt, math, PySimpleGUI as PSG, os
import urllib.parse
import data.cf as cf
from fillpdf import fillpdfs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#DATA LOADING
if __name__ == "__main__":
  PSG.theme("DarkBlack")

benzina = 1.95
today = dt.date.today()
current_year = today.year
current_month = today.month
current_day = today.day
current_dir, _ = str(os.path.realpath(__file__)).replace("\\", "/").rsplit("/", 1)

with open(f"{current_dir}/data/JSON/dt.json", "r", encoding="utf-8") as f:
  dit = json.load(f)

home_tab = [
  [PSG.Text("Nome Gara", s=(10,1)), PSG.Input("Nome in locandina", key="-NOME-GARA-", s=(50,1))], 
  [PSG.Text("Tipo Gara", s=(10,1)), PSG.Combo(["REG", "INTREG", "NAZ"], "Tipo", s=(8,1), key="-TIPO-GARA-", button_background_color="gray", button_arrow_color="white", enable_events=True)],
  [PSG.Text("Città Gara", s=(10,1)), PSG.Input("Città", key="-LUOGO-GARA-", size=(50,1))], 
  [PSG.Text("Indirizzo Gara", s=(10,1)), PSG.Input("Via", key="-INDIRIZZO-GARA-", size=(50,1))], 
  [PSG.Text("Data Gara", s=(10,1)), 
  PSG.Combo(["%02d" % x for x in range(1, 32)], default_value="Giorno", key="-GIORNO-GARA-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
  PSG.Text("/"),
  PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="Mese", key="-MESE-GARA-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
  PSG.Text("/"),
  PSG.Combo([x for x in range(current_year - 1, current_year + 2)][::-1], default_value="Anno", key="-ANNO-GARA-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
  [PSG.Text("Convocazione", s=(10,1)), 
  PSG.Combo(["%02d" % x for x in range(1, 32)], default_value="Giorno", key="-GIORNO-CONV-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
  PSG.Text("/"),
  PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="Mese", key="-MESE-CONV-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
  PSG.Text("/"),
  PSG.Combo([x for x in range(current_year - 1, current_year + 2)][::-1], default_value="Anno", key="-ANNO-CONV-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
  [PSG.Text("Data firma", s=(10,1)), 
  PSG.Combo(["%02d" % x for x in range(1, 32)], default_value="%02d" % current_day, key="-GIORNO-FIRMA-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
  PSG.Text("/"),
  PSG.Combo(["%02d" % x for x in range(1, 13)], default_value="%02d" % current_month, key="-MESE-FIRMA-", s=(8,1), button_background_color="gray", button_arrow_color="white"),
  PSG.Text("/"),
  PSG.Combo([x for x in range(current_year - 1, current_year + 2)][::-1], default_value=current_year, key="-ANNO-FIRMA-", s=(8,1), button_background_color="gray", button_arrow_color="white")],
  [PSG.Button("Export", key="-EXPORT-", disabled=True, bind_return_key=True, button_color="gray")],
  [PSG.Push()],
  [PSG.Text("Output")],
  [PSG.Multiline(disabled=True, no_scrollbar=True, autoscroll=True, expand_x=True, auto_refresh=True, size=(1, 5), key="-OUTPUT-TERMINAL-")],
  [PSG.Button("Clear output", key="-CLR-OUT-", button_color="gray")]
]
dit_list = ([
   PSG.Checkbox(text="", key=f"-CONVOCATO-{person["NumFIS"]}-", s=(1,1)), 
   PSG.Text(f"{person["Cognome"]} {person["Nome"]}",s=(34,1)),
   PSG.Input("", key= f"-GIORNI-{person["NumFIS"]}-", s=(5,1)),
   PSG.Checkbox(text="", key=f"-EXTRA-{person["NumFIS"]}-", s=(1,1), pad=(15,0))] for person in dit)
peeps_tab = [
  [PSG.Text("Conv."), PSG.Push(), PSG.Text("Arbitro"), PSG.Push(), PSG.Text("Giorni "), PSG.Text("Extra"), PSG.Text("   ")],
  [PSG.Column(dit_list, s=(1,300), vertical_scroll_only=True, expand_x=True, scrollable=True, sbar_arrow_color="white", sbar_background_color="grey")]
]
default_view = [
  [PSG.TabGroup(
      [
        [PSG.Tab("Home", home_tab)],
        [PSG.Tab("Arbitri", peeps_tab)]
      ]
    )
  ]
]

def get_distance(start: list[str]|str, destination:str):
  
  #URL Encode
  url_start:list[str] = []
  if type(start) == str:
    url_start.append(urllib.parse.quote_plus(start))
  elif type(start) == list:
    for place in start: url_start.append(urllib.parse.quote_plus(place))
  else:
   raise TypeError(f"Invalid Type for get_distance. Use list or str not {type(start)}")
  
  url_destination:str = urllib.parse.quote_plus(destination)
  
  #Chrome Setup
  options = webdriver.ChromeOptions()
  options.add_argument("--disable-gpu")
  options.add_argument("--disable-extensions")
  options.add_argument("--disable-search-engine-choice-screen")
  driver = webdriver.Chrome(options=options)
  wait = WebDriverWait(driver, 10)
  
  #Accept cookies
  driver.get("https://maps.google.com/")
  driver.find_element(By.CLASS_NAME, "lssxud").click()

  distances:dict[str, str] = {}
  for place in url_start:
    try:
      driver.get(f"https://www.google.com/maps/dir/{place}/{url_destination}")
      wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@aria-label='Auto']"))) #Wait for page to finish loading
      driver.find_element(By.XPATH, "//img[@aria-label='Auto']").click()
      wait.until(EC.text_to_be_present_in_element((By.CLASS_NAME, "XdKEzd"), "")) #Wait for page to finish loading
      time, route = driver.find_element(By.CLASS_NAME, "XdKEzd").text.split("\n") #Get best routes
      distance, unit = route.split(" ")
      place = urllib.parse.unquote_plus(place)
      distances[place.upper().replace("+", " ")] = float(distance.replace(",", "."))
    except Exception as e:
      print(f"Exception encountered while fetching distance from {place}: {e}")
      place = urllib.parse.unquote_plus(place)
      distances[place.upper()] = 0
  driver.quit()
  return distances

if __name__ == "__main__":

  window = PSG.Window(f"Rimborsi Arbitri | by Piombo Andrea", default_view, finalize=True)
  
  while True:

    events, values = window.read()
    if events == PSG.WIN_CLOSED: break
    if events == "-TIPO-GARA-": window["-EXPORT-"].update(disabled = False)
    if events == "-CLR-OUT-": window["-OUTPUT-TERMINAL-"].update("")

    #Aggiornamento dati sui pagamenti, gettoni.json: dati sui rimborsi; viaggi.json: distanze in KM dal luogo di gara
    with open(f"{current_dir}/data/JSON/gettoni.json", "r") as f:
      rimborsi = json.load(f)
    viaggi = get_distance(..., values["-INDIRIZZO-GARA-"])
    form_fields = list(fillpdfs.get_form_fields(f"{current_dir}/data/template.pdf").keys())

    if events == "-EXPORT-":
      
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
            datanascita, _ = person["DataNascita"].split("T")
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
            datarinnovo, _ = person["DataRinnovo"].split("T")
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