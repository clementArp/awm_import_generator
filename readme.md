# AWM Import Generator

Outil permettant de **générer des fichiers d'import pour les applications AWM**.

Les scripts principaux se trouvent dans le dossier :

```
src/
```

Exemples :

```
src/set_diag_app.py
src/set_prod_app.py
```

Les fichiers générés sont enregistrés dans le dossier :

```
out/
```

---

# Prérequis

- **Python 3.9**

Vérifier la version installée :

```bash
python --version
```

---

# Installation

## 1️⃣ Créer un environnement virtuel

Depuis la racine du projet :

```bash
py -3.9 -m venv .venv
```

---

## 2️⃣ Activer l'environnement virtuel

Sous Windows :

```bash
.venv\Scripts\activate
```

Le terminal doit afficher :

```
(.venv)
```

---

## 3️⃣ Installer les dépendances

```bash
pip install -r requirements.txt
```

---

# Utilisation

## Générer un import diagnostic

```bash
py .\src\set_diag_app.py
```

---

## Générer un import production

```bash
py .\src\set_prod_app.py
```

---

# Résultat

Les fichiers générés seront disponibles dans le dossier :

```
out/
```

---

# Structure du projet

```text
awm_import_generator
│
├─ src/
│   ├─ set_diag_app.py
│   └─ set_prod_app.py
│
├─ out/
│
├─ requirements.txt
└─ README.md
```

---

# Bonnes pratiques

Avant de lancer un script :

- activer l'environnement virtuel
- vérifier que les dépendances sont installées
- vérifier que le dossier `out/` existe (sinon il sera créé automatiquement)
