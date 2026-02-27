Excel Comparison Tool

Outil léger (Gradio + Python) pour comparer deux fichiers Excel sur une colonne clé et générer un fichier Excel de sortie avec les différences colorées.


Compare deux feuilles Excel en utilisant une colonne clé (gère les clés dupliquées).

Normalise les colonnes (accents, casse), nettoie les valeurs et compare avec tolérance numérique.

Exporte un fichier .xlsx où les cellules différentes sont surlignées en rouge et les colonnes sont auto-ajustées.

Quick start
# cloner (exemple)
git clone https://github.com/TON-UTILISATEUR/excel-comparison-tool.git
cd excel-comparison-tool

# (recommandé) créer et activer un venv
python -m venv venv
# mac/linux
source venv/bin/activate
# windows
venv\Scripts\activate

# installer dépendances
pip install -r requirements.txt

# lancer l'app (ouvre le navigateur)
python comparaison_final_tool.py
Utilisation (résumé)

Charger les deux fichiers Excel.

Choisir l’onglet et la ligne d’en-tête pour chaque fichier.

Choisir la colonne clé (bouton « Charger colonnes »).

Cliquer sur Comparer et exporter Excel → récupérer le fichier généré.

Fonctionnalités clés

Gestion des occurrences multiples via un index occurrence.

safe_compare() pour comparaison sûre (texte + numérique avec arrondi).

Analyse « Top 5 » des colonnes les plus uniques.

Export .xlsx avec cellules en rouge pour les différences et largeur de colonnes automatique.
