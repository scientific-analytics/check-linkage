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

- `DocLinkage` : **348 lignes** = 333 lignes de Step 3 + 15 linkages manuelles, aucun `Link` en double.
- `Problems` : **0 groupe, 0 ligne** — tous les groupes de `linkage_problems.csv` ont été tranchés (session du 2026-10-22).
- `Decisions` : **40 lignes**.

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
| 19 | Alphabet | 5093:DOB → 2788:VUH (Class A), share 291.328, 33.582 et 3723:KAH → 3742:JAS (Class C), share 345.504, 32.272, 2015-10-06 ; réorganisation Google → Alphabet 1 pour 1 (décision changée le 2026-11-03, avant : no linkage) |
| 56 | Biglari | 6196:BAN → 5371:PYR (Class B), 2018-05-01, 242.69, share 2.068 ; 1553:TYQ Class A = new security |
| 143 | Fidelity National Financial | 1094:QIW → 2291:FYR (FNF), 2014-06-30, 32.76, share 276 ; 3806:KIF FNFV = new security |
| 157 | Lundbeck | ~~1119:MAK → 3414:HIW (B Shares), 2022-06-10~~ **linkage retirée le 2026-11-04** : éclatement en actions A et B (split CIQ de 5), valeur répartie sur deux classes ; 6922:SIK A Shares = new security |
| 129 | Esprit Energy Trust | no linkage : facteur d'ajustement (Common Shares échangées ~4 pour 1 part + actions ProspEx le 2004-10-05, non ajusté dans CIQ) |
| 189 | Liberty Global | no linkage : 2 old (Series A/C) / 2 new (Class A/C), échange 1 pour 1 Inc → plc le 2013-06-07 ; le ratio 2:1 des prix vient de CIQ (nouveaux titres ajustés du split 2:1 du 2014-03-04, pas les anciens) : une linkage créerait un saut de −51 % ; les nouveaux titres ont leur propre historique depuis 2004/2005 (commentaire corrigé le 2026-11-03) |
| 193 | Lionsgate | no linkage : merge Class A + Class B → Common Shares le 2025-05-06, 2 old / 1 new, ratio non-1:1 |
| 194 | Lionsgate | no linkage : split Old Common Shares → Class A + Class B le 2016-12-09, 1 old / 2 new, ratio non-1:1 |
| 204 | Mediclinic | no linkage : fusion avec Al Noor Hospitals (2016-02-16), changement de devise ZAR→GBP et de bourse (Johannesburg→Londres), ratio non-1:1 |
| 230 | Odfjell | no linkage : Series A et Series B sont deux classes permanentes cotées depuis 1992, toujours cotées, prix différents (A 125 / B 111 NOK au 2005-06-17) ; AUC alterne entre elles (les deux en 2005-03 et 2005-06, B 2005-09/2005-12, A 2006-03, B 2008-03) ; société dans 6 reviews sur 98, pas d'événement corporate ; chaque classe a son historique CIQ depuis 1992 |
| 240 | Paramount Skydance | no linkage : fusion Skydance (2025-08-06), consolidation Class A + Class B → Class B, ratio non-uniforme (A 1.5:1, B 1:1), pas d'overlap AUC |
| 241 | Paramount Skydance | no linkage : spin-off Viacom/CBS (2006-01-03), Old Viacom scindé en CBS Corporation (ce companyid) et New Viacom Inc. (companyid 25403751, non suivi dans ce pipeline), ratio actions non uniforme (0.47 B, 0.50 A) |
| 287 | SunPower | 3897:BOD (Class A, la plus liquide) → 7045:MEM (Common Stock), 2011-11-17, 7.50, share 1 ; reclassification 1-pour-1 confirmée par SEC/Jones Day ; 7965:KUG Class B non lié |
| 288 | Sunshine Lake Pharma | no linkage : gap de 7 jours ouvrés et saut de prix ~4.3x entre 3095:BEG (13.02) et 2966:WUC (56.25), trop important |
| 299 | Telecom Italia | no linkage : reverse merger avec la société mère Olivetti S.p.A. (2003-08-04), prix/actions incohérents entre Old Ordinary Shares et les 2 nouvelles lignes (Ordinary + Saving Shares) |
| 341 | Warner Bros. Discovery | ~~7378:NOD → 7463:MON, 2022-04-08~~ **linkage retirée le 2026-11-04** : fusion Discovery / WarnerMedia avec émission d'actions (share 170 → 2427) ; 6251:CEV Series C non liée |
| 342 | Warner Bros. Discovery | no linkage : restructuration Discovery Holding Company → Discovery Communications (2008-09-18), chaque action convertie en 0.5 Series A + 0.5 Series C (SEC S-4 confirmé), ratio non-1:1 |
| 351 | Wilh. Wilhelmsen | no linkage : Class A et Class B sont deux classes permanentes cotées depuis 1992, toujours cotées, prix différents ; AUC alterne entre elles (A 2013-03/2014-03, B 2014-06, A 2015-06/2018-03) et les a toutes les deux en 2017 et depuis 2024-12 ; société seulement sporadiquement dans AUC (12 reviews sur 98, trous de ~7 ans et ~6.75 ans), pas d'événement corporate ; chaque classe a son historique CIQ depuis 1992 |

## Revue des sauts de prix (session du 2026-11-04)

Revue des linkages de DocLinkage v5 dont le saut de prix (`jump_pct`, dernier prix de l'ancien titre → premier
prix du nouveau) est supérieur à 2 % en valeur absolue. Pour chaque linkage : séries CIQ des deux trading items,
facteurs d'ajustement, splits, distributions, évolution du share, et recherche de l'opération (communiqués, SEC,
RNS, SENS). La tranche 2–5 % a été vérifiée avec sources ; le détail est dans les commentaires de
`input/manual_decisions.csv`.

### Règle

Une linkage veut dire que les deux séries de prix se recollent, facteurs d'ajustement CIQ appliqués, sans saut.

- **Garder** : échange 1:1 pour les porteurs de l'ancien titre, sans cash, sans scission, sans nouvelle classe et
  sans émission d'actions pour une autre société. Cas typiques : nouvelle holding, changement de domicile,
  conversion de trust ou de LP en société, transfert de cotation, conversion 1:1 d'une classe en classe unique,
  actionnaire de contrôle qui distribue sa participation (changement de flottant), regroupement pur couvert
  par le facteur CIQ (Breedon).
- **Retirer** : fusion avec émission d'actions (même 1:1 pour nos porteurs), SPAC, cash versé aux porteurs, ratio
  différent de 1:1, retour de capital (B shares), scission ou distribution d'une autre société, sortie de faillite,
  éclatement en plusieurs classes. Dans le doute, retirer (LondonMetric).

### Linkages retirées (59)

Toutes en « linkage removed » dans `input/manual_decisions.csv` (decided_on 2026-11-04).

| Saut | Groupes retirés |
|---|---|
| 10–20 % | 142 Fibria, 283 Sprint, 118 Designer Brands, 353 Williams Partners, 301 Berkeley, 359 Zillow, 126 Enviri, 266 Saxon Capital, 286 Stride, 41 BGC |
| 5–10 % | 32 Apartment Income REIT, 68 Brookfield Property REIT (GGP 2018), 177 Johnson Controls, 271 Seitel, 327 Vermilion, 157 Lundbeck, 361 la Rinascente, 182 Kirkland Lake Gold, 167 IHS Markit, 42 Banc of California, 267 Scentre, 341 Warner Bros. Discovery, 92 ChampionX, 339 WTW, 231 Old API Wind-down (Aralez), 43 Banco BPM, 198 LondonMetric (dans le doute) |
| 2–5 % | 268 Seabridge (+ ligne du fichier livré), 274 Sika, 356 Xstelos, 357 YTB, 200 Marshalls, 2 ADTRAN (+ ligne du fichier livré), 23 Altice, 113 DSM-Firmenich, 192 Lion Finance, 209 NEPI Rockcastle, 278 Six Flags, 362 nCino, 5 Xerium/ANDRITZ (+ ligne du fichier livré), 31 Antero Midstream, 58 BlackRock, 66 Broadcom, 127 Envision, 148 GCI Liberty, 154 Gulf Development, 255 R1 RCM, 280 Spectrum Brands, 202 Mebuki 2016, 291 TCF, 329 Viatris, 345 WestRock, 20 Alta Mesa, 69 GGP 2010, 96 Argo (Clearbrook), 150 Gramercy, 216 Noble, 276 Sirius XM 2024, 350 Whole Earth |

### Linkages gardées après examen (mouvement de marché)

- 10–20 % : Sibanye, Advantage Energy, Coast Entertainment.
- 5–10 % : APA, Lazard, LIN Media, Criteo, Bank of Ireland, Mebuki 2008, Centamin, Northam, Clear Channel Outdoor.
- 2–5 % (43) : notamment Suzano, CBD et SunPower (classe unique 1:1), Aroundtown (transfert de cotation),
  Veren (conversion de trust via Wild River, 1:1), Breedon (regroupement 5:1 couvert par CIQ), Metlen (option
  cash du squeeze-out ignorée), Tronox (rachat de Cristal traité comme opération distincte), Black Knight et
  Calumet (à revoir si la règle sur l'émission est appliquée strictement), Aimia, Trinidad, Boyd, Iyogin et
  Growthpoint (1:1 confirmé seulement par des sources secondaires et CIQ).

Résultat : DocLinkage v5 passe de 246 à 187 linkages.

### Changement de `Step_4b_rules.py`

Une ligne « linkage removed » **sans groupe** dans `manual_decisions.csv` retire la ligne du fichier livré
(`input/Historical_DocLinkage.xlsx`) qui a le même couple (assetid, Link). Utilisé pour Seabridge, ADTRAN et
ANDRITZ Fabrics.

### Points ouverts de cette revue

1. Linkages avec un saut < 2 % (132) non revues une par une.
2. Lundbeck (157) et Warner Bros. Discovery (341) apparaissent à la fois en « manual linkage added » et en
   « linkage removed » dans Decisions : supprimer les anciennes lignes « manual linkage added ».
3. La règle « price jump » de Step 4b compare des `priceclose` sans facteurs d'ajustement (ex. Berkeley +11.6 %
   = facteur 0.891).
4. Devises manquantes dans DocLinkage : Criteo `new_currency` = USD et Rheinmetall `old_currency` = EUR (CIQ),
   les trading items sont absents de `auc.csv`.
5. Données AUC : share des trust units de Crescent Point erratique en 2008–2009 (113 / 19 / 125 / 20), assets
   d'anciens titres non fermés dans `auc_asset.csv` (`until` = 2099-12-31, ex. 6362:HEP, 5372:DUP ACMP, 3723:LIC).

## Groupes restant dans `Problems`

Aucun — tous les groupes de `linkage_problems.csv` ont été tranchés (session du 2026-10-22). Voir la table
ci-dessus pour le détail des décisions.

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
