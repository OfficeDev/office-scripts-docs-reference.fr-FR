# <a name="office-scripts-api-documentation-tools"></a>Office Outils de documentation de l’API scripts

Ces outils vous aident à Office documentation SCripts et à l’équipe derrière elle. Suivez ces instructions pour exécuter les outils dans ce dossier.

## <a name="coverage-tester"></a>testeur de couverture

Cet outil fournit une vue d’ensemble de la couverture de la documentation pour chaque API. Chaque API est évaluée pour la qualité de la documentation et la présence d’exemples de code. Les mesures de qualité sont toujours en cours de développement.

La sortie de cet outil est un `.csv` fichier.

### <a name="coverage-tester-instructions"></a>Instructions du testeur de couverture

1. Clonez ou bifurcationz le repo.
1. Dans une fenêtre de commande, allez à `/office-scripts-docs-reference/generate-docs/tools`
1. Exécuter `npm install`
1. Exécuter `npm run build`
1. Exécuter `node coverage-tester`
1. Ouvrez « Couverture d’API Report.csv »
