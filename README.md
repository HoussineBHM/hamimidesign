# 🏦 Bank → Odoo Converter

Convertit automatiquement les relevés bancaires en fichier d'import Odoo.

## Fonctionnalités
- Upload multi-fichiers (CSV, XLS, XLSX)
- Détection automatique du format source
- Débit → négatif / Crédit → positif
- Date au format `yyyy-mm-dd`
- Export avec une feuille par mois

## Formats supportés
| Banque | Format | Exemple |
|--------|--------|---------|
| CFG Bank MAD | XLS multi-feuilles | `CFG_MAD.xls` |
| CFG Bank Devise | XLS mono-feuille | `CFG_CONV.xls` |
| AJW MAD/Devise | CSV | `AJW_MAD_*.csv` |

## Lancer en local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Déployer sur Streamlit Cloud
1. Push ce repo sur GitHub
2. Aller sur [share.streamlit.io](https://share.streamlit.io)
3. Connecter le repo → Deploy

## Import dans Odoo
`Comptabilité → Relevés bancaires → Importer`
