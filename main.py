"""
===============================================================================
Description :
-------------
Ce script permet de transformer un fichier Excel contenant plusieurs lignes
de contacts par société en un fichier CSV ne contenant qu'une seule ligne par
société.

Les informations de la société sont conservées une seule fois, tandis que
les informations de contact sont "aplaties" sur la même ligne, sous forme de
blocs répétés et suffixés (_1, _2, _3, etc.), en fonction du nombre maximum
de contacts trouvés pour une même raison sociale.

Structure attendue du fichier Excel (colonnes obligatoires) :
--------------------------------------------------------------
A  - Raison Sociale
B  - Description d'activité
C  - Population
D  - Localisation (departement)
E  - Civilité
F  - Nom
G  - Prénom
H  - Fonction
I  - Email pro
J  - Localisation
K  - Date de Création
L  - CEO
M  - SIREN
N  - SIRET
O  - Descriptif NAF

Fonctionnement :
----------------
- Lecture du premier onglet du fichier Excel
- Regroupement des lignes par "Raison Sociale"
- Détermination du nombre maximum de contacts par société
- Génération dynamique des colonnes de contacts suffixées
- Écriture du résultat dans un fichier CSV encodé en UTF-8 (BOM)
  avec un séparateur ';'

Entrée :
--------
- Fichier Excel (.xlsx)

Sortie :
--------
- Fichier CSV (.csv) avec une ligne par société

Utilisation :
-------------
python aplatir_contacts_par_societe.py <fichier_input.xlsx> <fichier_output.csv>
===============================================================================
"""

import sys
import pandas as pd

def aplatir_contacts_par_societe(path_excel: str, path_csv_sortie: str):

    df = pd.read_excel(path_excel, sheet_name=0, dtype=str)
    df = df.fillna('')

    colonnes_societe = [
        'Raison Sociale',
        "Description d'activité",
        'Population',
        'Localisation (departement)'
    ]
    colonnes_contact = [
        'Civilité',
        'Nom',
        'Prénom',
        'Fonction',
        'Email pro',
        'Localisation',
        'Date de Création',
        'CEO',
        'SIREN',
        'SIRET',
        'Descriptif NAF'
    ]

    manquantes = [col for col in colonnes_societe + colonnes_contact if col not in df.columns]
    if manquantes:
        print(f"❌ Erreur : colonnes manquantes dans l'Excel d'entrée : {manquantes}")
        sys.exit(1)

    groupes = df.groupby('Raison Sociale', sort=False)

    nombre_max_contacts = 0
    for raison, sous_df in groupes:
        n_contacts = len(sous_df)
        if n_contacts > nombre_max_contacts:
            nombre_max_contacts = n_contacts

    en_tetes_sortie = []
    en_tetes_sortie.extend(colonnes_societe)
    for i in range(1, nombre_max_contacts + 1):
        for champ in colonnes_contact:
            en_tetes_sortie.append(f"{champ}_{i}")

    listes_lignes_sortie = []
    for raison, sous_df in groupes:
        premiere_ligne = sous_df.iloc[0]
        infos_societe = [premiere_ligne[col] for col in colonnes_societe]

        contacts_liste = []
        for _, row in sous_df.iterrows():
            ligne_contact = [row[col] for col in colonnes_contact]
            contacts_liste.append(ligne_contact)

        nb_contacts_courant = len(contacts_liste)
        if nb_contacts_courant < nombre_max_contacts:
            blocs_vides = nombre_max_contacts - nb_contacts_courant
            for _ in range(blocs_vides):
                contacts_liste.append([''] * len(colonnes_contact))

        ligne_complete = infos_societe
        for contact in contacts_liste:
            ligne_complete.extend(contact)

        listes_lignes_sortie.append(ligne_complete)

    df_sortie = pd.DataFrame(listes_lignes_sortie, columns=en_tetes_sortie)
    df_sortie.to_csv(path_csv_sortie, index=False, encoding='utf-8-sig', sep=';')
    print(f"✅ Fichier CSV généré avec succès : {path_csv_sortie}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage : python main.py <fichier_input.xlsx> <fichier_output.csv>")
        sys.exit(1)

    fichier_input = sys.argv[1]
    fichier_output = sys.argv[2]
    aplatir_contacts_par_societe(fichier_input, fichier_output)
