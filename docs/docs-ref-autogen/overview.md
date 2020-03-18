---
title: Référence de l’API des scripts Office
description: Vue d’ensemble des API JavaScript pour les scripts Office
ms.date: 12/12/2019
ms.openlocfilehash: b3e75cbecac13f41c83f019c681b0682571ead41
ms.sourcegitcommit: 2834ee43a14a4b0953d5fb22749c115a42960e20
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42704475"
---
# <a name="office-scripts-api-reference"></a>Référence de l’API des scripts Office

L’API de scripts Office vous permet d’automatiser des tâches courantes dans Excel sur le Web. Utilisez cette documentation de référence pour en savoir plus sur les classes, les méthodes et les autres types disponibles pour vos scripts. Tous les objets accessibles par le biais de scripts Office sont disponibles dans la table des matières sur la gauche de la page.

## <a name="common-classes"></a>Classes communes

La liste suivante décompose les concepts de base du modèle objet de scripts Office. Cela indique les classes communes et la façon dont elles sont liées les unes aux autres.

- Un [classeur](/javascript/api/office-scripts/excel/excel.workbook) contient une ou plusieurs [feuilles de calcul](/javascript/api/office-scripts/excel/excel.worksheet) dans un [WorksheetCollection](/javascript/api/office-scripts/excel/excel.worksheetcollection).
- Une [feuille de calcul](/javascript/api/office-scripts/excel/excel.worksheet) donne accès aux cellules par le biais d’objets [Range](/javascript/api/office-scripts/excel/excel.range) .
- Une [plage](/javascript/api/office-scripts/excel/excel.range) représente un groupe de cellules contiguës.
- Les [plages](/javascript/api/office-scripts/excel/excel.range) permettent de créer et de placer des [tableaux](/javascript/api/office-scripts/excel/excel.table), des [graphiques](/javascript/api/office-scripts/excel/excel.chart), des [formes](/javascript/api/office-scripts/excel/excel.shape)et d’autres objets d’organisation ou de visualisation de données.
- Une [feuille de calcul](/javascript/api/office-scripts/excel/excel.worksheet) contient des collections de ces objets de données (tels qu’un [ChartCollection](/javascript/api/office-scripts/excel/excel.chartcollection)) qui sont présents dans la feuille individuelle.
- [Classeurs](/javascript/api/office-scripts/excel/excel.workbook). contiennent des collections de certains de ces objets de données (par exemple, un [TableCollection](/javascript/api/office-scripts/excel/excel.tablecollection)) pour l’intégralité du [classeur](/javascript/api/office-scripts/excel/excel.workbook).

Pour plus d’informations sur le modèle objet de scripts Office, consultez la [rubrique Scripting Fundamentals for Office scripts in Excel on the Web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>Voir aussi

- [À propos des scripts Office](/office/dev/scripts/overview/excel)
- [Enregistrer, modifier et créer des scripts Office dans Excel sur le Web](/office/dev/scripts/tutorials/excel-tutorial)
- [Scripts de base pour les scripts Office dans Excel sur le Web](/office/dev/scripts/develop/scripting-fundamentals)
