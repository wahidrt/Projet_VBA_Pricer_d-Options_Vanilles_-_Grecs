# Historique des versions

## 1.1.1 — 25 juillet 2026

- restauration du dernier classeur `.xlsm` créé et validé par Microsoft Excel ;
- abandon de la réécriture directe du projet VBA binaire hors d’Excel ;
- ajout d’un installateur Windows générant `Projet_VBA_BS_corrige.xlsm` ;
- protection du classeur modèle : les corrections sont importées par Excel
  dans un nouveau fichier, sans modifier l’original ;
- ajout d’un contrôle d’intégrité du binaire Excel de référence.

## 1.1.0 — 25 juillet 2026

- correction de l’accumulation des résultats entre les points de simulation ;
- suppression sécurisée des positions sans déplacer le tableau de bord ;
- conversion systématique des maturités de jours en années ;
- validation complète des paramètres et des positions ;
- prise en compte du rendement de dividende continu ;
- contrôle mono-sous-jacent et maturité commune ;
- migration ciblée du sous-jacent incohérent dans le jeu d’exemple historique ;
- amélioration de la gestion des erreurs et du nettoyage d’Excel ;
- ajout d’unités explicites pour les Grecs ;
- amélioration et sécurisation des graphiques ;
- ajout de tests automatisés et de scripts de synchronisation VBA ;
- normalisation UTF-8 des sources et enrichissement de la documentation.
