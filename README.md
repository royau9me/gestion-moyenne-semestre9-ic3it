# 📊 Gestion Moyenne Semestre 9 IC3 IT/ESTP

<div align="center">

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![.NET](https://img.shields.io/badge/.NET-8.0-purple)
![License](https://img.shields.io/badge/license-MIT-green)

**Application Windows Forms pour la gestion et le calcul des moyennes du Semestre 9 IC3 IT/ESTP**

</div>

## 🎯 Fonctionnalités

### ✨ Calcul intelligent des moyennes
- ✅ Calcul automatique des moyennes pondérées par coefficient
- ✅ Validation des saisies (moyennes entre 0 et 20)
- ✅ Affichage coloré selon les résultats (vert, orange, rouge)
- ✅ Moyenne générale du semestre

### 📑 Gestion complète des matières

**UE de Culture Générale** (5 matières)
- Stage de production (Coef. 2)
- Séminaire des méthodes (Coef. 1)
- Choix économiques des projets (Coef. 1)
- E.P.S (Coef. 1)
- Anglais (Coef. 1)

**UE de Spécialités** (7 catégories)
- Sciences de la construction
- Techniques d'entretien des infrastructures
- Techniques et Méthodes de Contrôle et Gestion
- Gestion des infrastructures
- Hydraulique appliquée
- Conception des infrastructures
- Projet de fin d'études (PFE - Coef. 4)

### 💾 Exports multiples
- **📄 PDF** : Document formaté avec tableaux professionnels
- **📊 Excel** : Tableur avec formules et mise en forme
- **📝 TXT** : Export texte simple

## 🖥️ Technologies

- **Framework** : .NET 8.0 Windows Forms
- **Langage** : C# 12
- **Bibliothèques** :
  - `iText7 (8.0.4)` - Génération de PDF
  - `ClosedXML (0.104.1)` - Génération de fichiers Excel
  - `Newtonsoft.Json (13.0.3)` - Manipulation JSON

## 📦 Installation

### Pour les utilisateurs

1. Téléchargez l'installateur : `Setup_GestionMoyenne_IC3IT_v1.0.exe`
2. Exécutez l'installateur (aucune installation .NET requise)
3. Suivez les instructions
4. Lancez depuis le menu Démarrer ou le raccourci bureau

### Pour les développeurs
```bash
# Cloner le repository
git clone https://github.com/royau9me/gestion-moyenne-semestre9-ic3it.git

# Naviguer dans le dossier
cd gestion-moyenne-semestre9-ic3it

# Restaurer les packages NuGet
dotnet restore

# Compiler
dotnet build -c Release

# Publier
dotnet publish -c Release -r win-x64 --self-contained true
```

## 📖 Utilisation

1. **Saisie** : Entrez vos heures dédiées et moyennes
2. **Calcul** : Cliquez sur "Calculer"
3. **Export** : Choisissez PDF, Excel ou TXT

## 🎓 Contexte

Application développée pour les étudiants **IC3 IT** (Infrastructures et Conception des Constructions - Informatique et Technologies) de l'**ESTP**.

## 👨‍💻 Auteur

**Développé par un étudiant IC3 IT - ESTP**
- GitHub : [@royau9me](https://github.com/royau9me)

## 📄 Licence

MIT License - Libre d'utilisation

## 🙏 Remerciements

- Mes camarades de l'ESTP IC3 IT
- La communauté .NET

---

⭐ **N'hésitez pas à mettre une étoile si ce projet vous aide !**
