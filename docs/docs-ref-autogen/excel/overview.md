---
title: Référence de l'API Office Scripts
description: Vue d’ensemble des API JavaScript de scripts Office.
ms.date: 06/29/2020
ms.openlocfilehash: 7c4fe97ca35cfb442ebbf9db2e0b03b389185ae8
ms.sourcegitcommit: 9c4c4c213a203e58c55eb3d84d7d92fa527f3eb8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/01/2020
ms.locfileid: "45004730"
---
# <a name="office-scripts-api-reference"></a>Référence de l'API Office Scripts

L’API de scripts Office vous permet d’automatiser des tâches courantes dans Excel sur le Web. Utilisez cette documentation de référence pour en savoir plus sur les classes, les méthodes et les autres types disponibles pour vos scripts. Tous les objets accessibles par le biais de scripts Office sont disponibles dans la table des matières sur la gauche de la page.

> [!NOTE]
> Si vous recherchez les API JavaScript pour le développement de compléments Office, consultez la référence de l' [API JavaScript pour les compléments Office](/javascript/api/overview?view=excel-js-preview).

## <a name="common-classes"></a>Classes communes

La liste suivante décompose les concepts de base du modèle objet de scripts Office. Cela indique les classes communes et la façon dont elles sont liées les unes aux autres.

- Un [classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) contient une ou plusieurs [feuilles de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet).
- Une [feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet) donne accès à des cellules via [plage](/javascript/api/office-scripts/excelscript/excelscript.range) objets.
- Une [plage](/javascript/api/office-scripts/excelscript/excelscript.range) représente un groupe de cellules contiguës.
- Les [plages](/javascript/api/office-scripts/excelscript/excelscript.range) sont utilisées pour créer et placer des [tableaux](/javascript/api/office-scripts/excelscript/excelscript.table), des [graphiques](/javascript/api/office-scripts/excelscript/excelscript.chart), des [formes](/javascript/api/office-scripts/excelscript/excelscript.shape) et d’autres objets d’organisation ou de visualisation de données.
- Une [feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contient des tableaux remplis avec les objets qui sont présents dans la feuille individuelle.
- Un [classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) contient des tableaux de certains de ces objets de données pour l’ensemble du classeur.

Pour plus d’informations sur le modèle objet de scripts Office, consultez la [rubrique Scripting Fundamentals for Office scripts in Excel on the Web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>Voir aussi

- [À propos des scripts Office](/office/dev/scripts/overview/excel)
- [Enregistrer, modifier et créer des scripts Office dans Excel sur le web](/office/dev/scripts/tutorials/excel-tutorial)
- [Principes de base des scripts pour Office Scripts dans Excel sur le web](/office/dev/scripts/develop/scripting-fundamentals)
