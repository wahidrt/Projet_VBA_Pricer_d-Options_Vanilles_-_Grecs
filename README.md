# Pricer d’options vanilles et Grecs — Excel VBA

Application Excel/VBA de valorisation et d’analyse d’un portefeuille
mono-sous-jacent d’options européennes avec le modèle de Black–Scholes.

![Tableau de bord du pricer](docs/dashboard.png)

![Courbes des Grecs](docs/greeks.png)

## Fonctionnalités

- valorisation des Calls et Puts européens ;
- prise en compte d’un rendement de dividende continu ;
- positions acheteuses et vendeuses ;
- calcul du Delta, Gamma, Vega, Theta et Rho ;
- profil de PnL mark-to-market et à l’échéance ;
- formulaire VBA pour ajouter et supprimer des positions ;
- contrôles de validité des paramètres et des positions ;
- graphiques automatiques du PnL et des Grecs ;
- tests numériques de non-régression.

## Hypothèses du modèle

Le modèle suppose :

- un seul sous-jacent par portefeuille ;
- une maturité commune pour construire un PnL à l’échéance cohérent ;
- des options européennes ;
- un taux sans risque, une volatilité et un rendement de dividende constants ;
- des montants exprimés par unité d’option, sans multiplicateur de contrat ;
- l’absence de coûts de transaction.

Les paramètres de taux et de volatilité sont saisis sous forme décimale :
`0,25` correspond à `25 %`.

## Utilisation

1. Téléchargez `Projet_VBA_BS.xlsm`.
2. Ouvrez le fichier avec Microsoft Excel sur Windows.
3. Autorisez les macros uniquement après avoir vérifié leur provenance.
4. Renseignez le spot, le taux, la volatilité, le dividende et la plage de simulation.
5. Cliquez sur **Ouvrir Portfolio** pour gérer les positions.
6. Cliquez sur **Calculer & Tracer** pour recalculer les résultats.

Une prime vide est automatiquement remplacée par la prime théorique
Black–Scholes. Une prime égale à zéro est conservée comme une valeur explicite.

## Unités des Grecs

| Mesure | Unité affichée |
|---|---|
| Delta | variation pour une unité de sous-jacent |
| Gamma | variation du Delta pour une unité de sous-jacent |
| Vega | variation pour +1 point de volatilité |
| Theta | variation par jour calendaire |
| Rho | variation pour +1 point de taux |

## Structure du dépôt

```text
.
├── Projet_VBA_BS.xlsm
├── src/
│   ├── Classe1.cls
│   ├── DesignModule.bas
│   ├── Portfolio.bas
│   ├── frmPortfolio.frm
│   └── frmPortfolio.frx
├── scripts/
│   ├── Export-VbaSources.ps1
│   └── Sync-VbaProject.ps1
├── tests/
└── docs/
```

Le fichier `.xlsm` contient la version corrigée du projet VBA et permet
d’utiliser directement l’application. Les fichiers du dossier `src` sont les
mêmes sources en version lisible, UTF-8 et comparable dans Git.

## Synchroniser les sources avec le classeur

Après une modification des fichiers dans `src`, ouvrez PowerShell à la racine
du dépôt et exécutez :

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Sync-VbaProject.ps1
```

Le script crée une sauvegarde du classeur, importe les modules, applique les
validations, recalcule la grille et sauvegarde le fichier.

Excel doit autoriser temporairement l’option :
**Accès approuvé au modèle d’objet du projet VBA**. Désactivez-la après la
synchronisation.

Pour exporter le projet VBA du classeur vers des fichiers UTF-8 :

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Export-VbaSources.ps1
```

## Tests

Les tests utilisent uniquement la bibliothèque standard de Python :

```bash
python -m unittest discover -s tests -v
```

Ils contrôlent notamment :

- la parité Call–Put avec dividende ;
- le Delta et le Gamma par différences finies ;
- les valeurs de référence du portefeuille exemple ;
- la réinitialisation des totaux à chaque valeur du spot ;
- la conversion systématique des maturités en jours ;
- la suppression sécurisée des positions.

## Limites et extensions possibles

Le projet ne traite pas encore :

- les options américaines ;
- les smiles et surfaces de volatilité ;
- les portefeuilles multi-actifs ;
- les courbes de taux par maturité ;
- les multiplicateurs de contrat et frais de transaction ;
- les scénarios dynamiques dans le temps.

Ce logiciel est un projet pédagogique. Il ne constitue pas un conseil
financier ni un système de valorisation destiné à la production.
