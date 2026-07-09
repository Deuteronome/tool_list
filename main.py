import warnings
import pandas as pd


warnings.filterwarnings(
    "ignore",
    message="Print area cannot be set*"
)

warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported*"
)

df = pd.read_excel('./src/catalogueOutils.xlsx')
df = df.drop(columns=['type de contenu', 'étape parcours','Socle', 'Créé dans Up2C',
       'A jour dans Up2C', 'Besoin RH', 'Besoin ordi', 'besoin matériel','A jour dans Sharepoint',
       'compétences', 'Unnamed: 17'])
df = df.rename(columns={'Nom de l\'outil':'name', 'programmation':'planning', 'up2c-projet':'project', 'up2c-séquence':'sequence',
       'modalité':'mode', 'objectifs généraux':'target', 'chemin':'path'})

df.to_csv('./output/tools.csv')

