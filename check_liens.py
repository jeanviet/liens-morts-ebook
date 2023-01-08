import os
import requests
from bs4 import BeautifulSoup
import openpyxl

# Créer un nouveau fichier Excel
wb = openpyxl.Workbook()

# Sélectionner la feuille de travail active
sheet = wb.active

# Spécifier les titres des colonnes du fichier Excel
sheet.append(["Chapitre","Texte de lien", "URL", "Code HTTP"])

# Définir le chemin du répertoire contenant les pages web locales
local_path = "/Users/jeanviet/test/pages/"

# Récupérer la liste des fichiers HTML dans le répertoire
html_files = [f for f in os.listdir(local_path) if f.endswith(".xhtml" or ".htm" or ".html")]

# Pour chaque fichier HTML, récupérer le contenu et parser les liens
for html_file in html_files:
  with open(os.path.join(local_path, html_file), "r") as f:
    html = f.read()
  soup = BeautifulSoup(html, "html.parser")
  links = soup.find_all("a")

  # Pour chaque lien, cliquer dessus et afficher le code HTTP
  for link in links:
    href = link.get("href")
    text = link.text

    # Si le lien ne pointe pas vers une URL valide, passer au suivant
    if not href or href.startswith("#") or href.startswith("../") or href.startswith("mailto"):
      continue

        # Cliquer sur le lien et récupérer le code HTTP
    try:
      response = requests.get(href)
      status_code = response.status_code
    except requests.exceptions.ConnectionError:
      # Si une erreur de connexion se produit, passer au lien suivant et mettre le code HTTP à "page KO"
      print("Erreur de connexion, passage au lien suivant")
      status_code = "page KO"

    # Afficher le texte du lien, l'URL et le code HTTP
    print(f"Texte de lien: {text}")
    print(f"URL: {href}")
    print(f"Code HTTP: {status_code}")
    print()

    # Ajouter une ligne au fichier Excel avec les informations du lien et de la page
    sheet.append([html_file, text, href, status_code])

# Enregistrer le fichier Excel
wb.save("suivi_liens.xlsx")