# HANDOFF — Fiche sécurité « Fortes chaleurs » (Aéro-Club de Brive)

Branche de travail : `claude/pilot-heat-guidelines-opr5u5`

## Objectif
Créer une **fiche de consignes de sécurité sur une page A4** pour les pilotes
privés sur avions légers, concernant le **vol par forte chaleur**. Trois volets :
1. **Performances avion** (altitude-densité, distances de décollage, moteur)
2. **Santé du pilote** (déshydratation, coup de chaleur, fatigue)
3. **iPad / tablettes EFB** (surchauffe et panne en vol)

Style demandé : **sobre / pro**, lisible, aéré, mise en page une page.

## Fichiers déjà produits (poussés sur la branche)
- `consignes-chaleur-pilotes.html` — document source (mise en page 3 colonnes
  + bandeau d'alerte + bloc « 5 réflexes essentiels » + en-tête avec logo).
- `consignes-chaleur-pilotes.pdf` — PDF généré depuis le HTML.
- `consignes-chaleur-pilotes.txt` — version texte brut (mémo / contexte).

## Génération du PDF
Le PDF est généré avec **WeasyPrint** :
```bash
pip install weasyprint        # si absent
python3 -m weasyprint consignes-chaleur-pilotes.html consignes-chaleur-pilotes.pdf
```

## État actuel
- ✅ Contenu rédigé et sourcé (FFPLUM, FFA Mémo VFR 2025-2026, FAA P-8740-02,
  AOPA, SKYbrary, iPad Pilot News).
- ✅ Mention « ULM » retirée (fiche centrée avions légers).
- ✅ Mise en page lisible et aérée (version validée).
- ⏳ **RESTE À FAIRE : intégrer le vrai logo de l'Aéro-Club de Brive.**

## Ce qui reste — LE LOGO
L'en-tête HTML pointe vers une image nommée **`logo-aeroclub-brive.png`** :
```html
<div class="logo">
  <img src="logo-aeroclub-brive.png" alt="Aéro-Club de Brive">
</div>
```
Le fichier image n'a pas encore pu être fourni (les images collées dans le chat
de l'environnement « web » n'étaient pas accessibles sur le disque).

### À faire dans le nouvel environnement
1. Déposer / joindre le fichier logo (préférer la **version bleue sur fond blanc**)
   dans le dossier du dépôt, nommé `logo-aeroclub-brive.png`.
2. Régénérer le PDF (commande WeasyPrint ci-dessus).
3. Commiter + pousser `logo-aeroclub-brive.png`, le HTML et le PDF mis à jour.

## Phrase de reprise suggérée
> « Lis HANDOFF.md. Intègre le logo que je te joins comme
> `logo-aeroclub-brive.png`, puis régénère le PDF et pousse le tout. »
