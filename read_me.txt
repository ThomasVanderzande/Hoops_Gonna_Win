Ce répertoire recense les fichiers suivants :

NBA Teams.zip : Recense les fichiers Excel contenant les stats pour chaque équipe
NBA_App_v2.xlsx : Recense les équipes du jour et le nombre de matchs joué pour l'instant cette saison (au 31 Avril)
scraping_basketball_reference.py : Script pour récupérer les stats de chaque équipe du jour
xtrain_all_para.xlsx : Full Dataset (train + test) - Input
ytrain_extended.xlsx : Full Dataset (train + test) - Output
NBAI_v2.py : Script qui analyse les matchs du jour repertoriés dans NBA_App_v2.xlsx et calcule les probabilités pour chaque équipe.
             Génère automatiquement un rapport NBA_report_{date}.xlsx
Today.xlsx : Template utilisé par NBAI_v2.py pour générer le rapport NBA_report_{date}.xlsx


Pour tester :

- Mettre le document du répertoire dans un même dossier
- Dézipper NBA Teams.zip dans le dossier. Le zip peut être supprimé.
- NBAI_v2.py peut être lancé, un fichier de type NBA_report_{date}.xlsx devrait être généré dans le dossier avec les matchs analysés.
/!\ Si vous lancez le script deux fois d'affilée, l'ancien "Nba_report" sera écrasé

