# Assistant de manipulation des données Excel – **Helpis**

**Helpis** est une application Python qui permet de :

- ✅ Détecter des erreurs et incohérences dans les fichiers Excel (`.xls` / `.xlsx`)
- 📊 Appliquer des tests statistiques sur les données
- 🔧 Optimiser, formater et restructurer les fichiers Excel

---

## 🛠 Prérequis

- [Python 3.13](https://www.python.org/downloads/)
- [`cx_Freeze`](https://pypi.org/project/cx-Freeze/)

> ⚠️ Lors de l’installation de Python, pensez à **cocher l’option "Add Python to PATH"**

---

## 🚀 Installation

### 1. Cloner le dépôt

```bash
git clone https://github.com/mariuschartier/Helpis.git
cd Helpis
```

### 2. Construire l'application

```bash
python setup.py build
```

L'exécutable `Helpis.exe` sera disponible dans le dossier suivant :

```
build/exe.win-amd64-3.13/
```

---

## ❓ FAQ

### 🔹 Comment installer Python ?
Téléchargez-le depuis le site officiel :  
👉 [https://www.python.org/downloads/](https://www.python.org/downloads/)  
Lors de l’installation, cochez bien **"Add Python to PATH"**.

---

### 🔹 Comment installer cx_Freeze ?
Une fois Python installé, ouvrez un terminal et tapez :
```bash
pip install cx-Freeze
```

---

### 🔹 Pourquoi la conversion ou l’optimisation d’un fichier prend-elle du temps ?
Ces fonctionnalités peuvent être **lentes**. Par exemple :
- Un fichier contenant **220 000 cellules** ou pesant **890 Ko**
- Peut prendre jusqu’à **10 minutes** à être traité

---

### 🔹 Pourquoi l'application s'appelle *Helpis* ?
Le nom *Helpis* est une combinaison de :
- **"Help"** (aider, en anglais)
- **"Elpis"**, la déesse de l’Espoir dans la mythologie grecque

🕊️ Cette application a été conçue pour vous accompagner dans la manipulation de vos fichiers Excel, avec efficacité… et dans l'espoir de vous aider.

---

## 📝 Licence

Ce projet est sous licence libre.  
<!-- *À compléter selon les besoins (MIT, GPL, etc.).* -->

---

## 💬 Contact

Pour toute suggestion, amélioration ou bug, n’hésitez pas à ouvrir une issue sur le dépôt GitHub :  
👉 [https://github.com/mariuschartier/Helpis](https://github.com/mariuschartier/Helpis)

---
