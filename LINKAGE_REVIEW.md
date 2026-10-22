# Revue manuelle des linkages (session du 2026-10-21)

État de la revue des groupes mis de côté par Step 2 (`output/linkage_problems.csv`) et des décisions prises.
Le résultat est dans **`output/Historical_DocLinkage_v2.xlsx`**, régénéré par `Step_4_manual_linkages.py`.

## Pipeline

| Étape | Script | Sortie |
|---|---|---|
| 1 | `Step_1_download_prices_around_linkage.py` | `output/linkage_groups.csv`, `output/prices_around_linkage.csv` |
| 2 | `Step_2_build_linkage.py` | `output/proposed_linkage.csv`, `output/linkage_problems.csv` |
| 3 | `Step_3_fill_historical.py` | `output/Historical_DocLinkage_new.xlsx`, `output/linkages/`, `output/problems/` |
| **4** | **`Step_4_manual_linkages.py`** | **`output/Historical_DocLinkage_v2.xlsx`** + log `output/manual_linkages_<timestamp>.log` |
| outil | `download_auc_timeseries.py [companyid ...]` | `output/auc_timeseries/<companyid>_<company>/` |

### Step 4 : `Step_4_manual_linkages.py`

Part de `output/Historical_DocLinkage_new.xlsx` (sortie de Step 3, non modifiée) et écrit `output/Historical_DocLinkage_v2.xlsx` :

- **onglet `DocLinkage`** : lignes de Step 3, moins `REMOVED_LINKAGES`, plus `MANUAL_LINKAGES` ajoutées à la fin
  (`stamp` = `AUC v2`, `note` = `-`) ;
- **onglet `Problems`** : groupes de `linkage_problems.csv` non tranchés, une ligne par trading item
  (`role` old / new, `assetid` toujours un seul id), colonne `note` remplie depuis `problem_notes.csv` ;
- **onglet `Decisions`** : une ligne par décision forcée.

Dictionnaires à éditer dans le script pour une nouvelle décision, puis relancer `python Step_4_manual_linkages.py` :

| Dictionnaire | Effet | Décision dans l'onglet `Decisions` |
|---|---|---|
| `MANUAL_LINKAGES` | linkage ajoutée à la fin de `DocLinkage` | manual linkage added |
| `NEW_SECURITIES` | nouvelle ligne d'un groupe sans linkage | new security, no linkage |
| `NO_LINKAGE` | groupe retiré de `Problems`, pas de linkage | no linkage |
| `NOT_VALIDATED` | circulaire refusée, ni dans `DocLinkage` ni dans `Problems` | circular linkage not validated |
| `REMOVED_LINKAGES` | linkage proposée par Step 2 retirée de `DocLinkage` | linkage removed |
| `CIRCULAR_PROBLEMS` | circulaire laissée dans `Problems` avec sa raison (vide aujourd'hui) | circular linkage left aside |

`problem_notes.csv` (assetid, Link, company, review, note) : notes de la revue manuelle. Ligne old = note de la
linkage qui part d'elle, ligne new = note de la linkage qui arrive sur elle. Les notes non utilisées sont signalées
dans le log (Esprit et Liberty Class B aujourd'hui).

### `download_auc_timeseries.py`

Pour chaque société (`COMPANIES` ou ids en argument) : périodes AUC lues dans
`/home/rudy/PycharmProjects/check_referential/data_gov/<date>/_universe/auc.csv` (une ligne est dans AUC de sa
review jusqu'à la review suivante), puis prix CIQ :

- `auc_timeseries.csv` / `auc_timeseries_long.csv` : prix téléchargés **uniquement sur les périodes AUC** ;
- `full_timeseries.csv` / `full_timeseries_long.csv` : historique CIQ complet ;
- `auc_periods.csv` : périodes de présence dans AUC.

Copies dans `output/problems/<group>_<company>/` pour Centamin (91), Coast (98), Cape (85), Man Group (199),
Old Mutual (233), Swiss Re (290), WCT (334), Esprit (129).

## État de `Historical_DocLinkage_v2.xlsx`

- `DocLinkage` : **346 lignes** = 333 lignes de Step 3 + 13 linkages manuelles, aucun `Link` en double.
- `Problems` : **13 groupes, 37 lignes**.
- `Decisions` : **22 lignes**.

Vérifié : toutes les linkages validées sont présentes avec les bonnes valeurs, toutes les non validées sont absentes,
les 333 lignes de Step 3 sont identiques et dans le même ordre.

## Linkages circulaires (review AUC du 2008-12-19)

À la review du **2008-12-19**, AUC a remplacé la ligne principale de 12 sociétés par une autre ligne
(Old Ordinary, Stapled...). La ligne principale revient plus tard : la linkage retour est en général déjà proposée
par Step 2. La linkage aller (2008) avait été écartée par Step 2 parce que la date estimée était fausse.

| Groupe | Société | Linkage aller | Prix | Date | Retour dans DocLinkage | Décision |
|---|---|---|---|---|---|---|
| 26 | American Tower | 3928:FOR → 1208:VEP | 32.61 | 2008-11-03 | 2012-01-04 | validée |
| 40 | AusNet | 1269:HYB → 3888:HAS | 1.08 | 2008-11-03 | 2015-06-09 | validée |
| 46 | Bank of Ireland | 7410:WEV → 1329:QAS | 63 | 2008-11-03 | 2017-07-10 | validée |
| 50 | Beazley | 1534:SUP → 6540:QIP | 1.06 | 2008-11-03 | 2009-06-09 et 2016-04-14 | validée |
| 85 | Cape | 1422:JYM → 5568:BUJ | 0.703 | 2008-11-03 | 2011-06-27 | validée (lignes AIM / Main Market, prix légèrement différents) |
| 91 | Centamin | 3953:NEG → 7339:QUQ | 0.29 | 2008-11-03 | 2011-12-30 | validée |
| 98 | Coast Entertainment | 3972:PAV → 1542:VIQ | 1.25 | 2008-11-03 | 2018-11-30 | validée |
| 199 | Man Group | 1380:CEF → 4561:TEM | 3.608 | 2008-11-03 | **aucun** | validée |
| 203 | Mebuki | 4522:PIS → 2701:VYT | 411.966 | 2008-11-03 | 2016-10-03 | validée |
| 290 | Swiss Re | 6325:WAB → 5431:JYJ | **49.16** | **2008-12-19** | **aucun** | validée, à la date de la bascule AUC |
| 233 | Old Mutual | 6441:MUT → 1247:WOG | 0.661 | 2008-11-03 | retour 2018-06-26 **retiré** | **non validée** : Johannesburg ZAR vs Londres GBP |
| 334 | WCT | 6109:HAJ → 6327:BYK | 1.209 | 2008-11-03 | aucun (WCT ne revient pas dans AUC) | **non validée** |

## Autres décisions

| Groupe | Société | Décision |
|---|---|---|
| 19 | Alphabet | no linkage (2788:VUH Class A, 3742:JAS Class C) |
| 56 | Biglari | 6196:BAN → 5371:PYR (Class B), 2018-05-01, 242.69, share 2.068 ; 1553:TYQ Class A = new security |
| 143 | Fidelity National Financial | 1094:QIW → 2291:FYR (FNF), 2014-06-30, 32.76, share 276 ; 3806:KIF FNFV = new security |
| 157 | Lundbeck | 1119:MAK → 3414:HIW (B Shares), 2022-06-10, 32.50, share 794.272 ; 6922:SIK A Shares = new security |
| 129 | Esprit Energy Trust | no linkage : facteur d'ajustement (Common Shares échangées ~4 pour 1 part + actions ProspEx le 2004-10-05, non ajusté dans CIQ) |

## Groupes restant dans `Problems`

Liberty Global (189), Lionsgate (193, 194), Mediclinic (204), Odfjell (230), Paramount Skydance (240, 241),
SunPower (287), Sunshine Lake Pharma (288), Telecom Italia (299), Warner Bros. Discovery (341, 342),
Wilh. Wilhelmsen (351). Les notes de la revue sont dans la colonne `note`.

## Points ouverts

1. **Date des 9 autres circulaires validées** : elles sont au 2008-11-03 (prix de la liste fournie), alors que la
   bascule AUC est au 2008-12-19 (comme Swiss Re). Prix de la nouvelle ligne au 2008-12-19 :
   American Tower 29.45, AusNet 0.895, Bank of Ireland 20.25, Beazley 1.28, Cape 0.535, Centamin 0.4125,
   Coast 0.92, Man Group 2.35, Mebuki 412.82. À décider.
2. **Linkages retour manquantes** : Man Group (4561:TEM → 1380:CEF, proposé 2012-11-12 à 0.8155) et Swiss Re
   (5431:JYJ → 6325:WAB, 2011-05-23 à 49.80 ou 2011-12-08 à 48.60). Share à trouver.
3. **Swiss Re 2011** : les deux lignes sont dans AUC aux reviews 2011-09-16 et 2011-12-16 (doublon sous le même asset).
4. **Bug potentiel Step 2** : si `Historical_DocLinkage_v2.xlsx` sert de `Historical_DocLinkage.xlsx`, les boucles
   A → B → A des circulaires font tourner `resolve_root_assets` sans fin (`while a in root`). Ajouter une garde
   avant de relancer Step 2.
5. Colonne `review` de `problem_notes.csv` (statuts « Not the same price », « ?? »...) non reportée dans l'Excel.
6. `data_gov/<date>/timeseries/<année>/prices.parquet` ne contient pas les trading items de Cape ni de Centamin
   (recherche par tradingitemid et securityid) : non expliqué.
