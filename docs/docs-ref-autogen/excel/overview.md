---
title: Référence de l'API Office Scripts
description: Vue d’ensemble des API JavaScript de scripts Office.
ms.date: 06/17/2020
ms.openlocfilehash: 5634d0e5f68464655054ad1c09eb7931e0da62d4
ms.sourcegitcommit: 163b26a43411ad7f13a01237efe9b8d6de656b47
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44882170"
---
# <a name="office-scripts-api-reference"></a>Référence de l'API Office Scripts

L’API de scripts Office vous permet d’automatiser des tâches courantes dans Excel sur le Web. Utilisez cette documentation de référence pour en savoir plus sur les classes, les méthodes et les autres types disponibles pour vos scripts. Tous les objets accessibles par le biais de scripts Office sont disponibles dans la table des matières sur la gauche de la page.

> [!NOTE]
> Si vous recherchez les API JavaScript pour le développement de compléments Office, consultez la référence de l' [API JavaScript pour les compléments Office](/javascript/api/overview?view=excel-js-preview).

## <a name="common-classes"></a>Classes communes

La liste suivante décompose les concepts de base du modèle objet de scripts Office. Cela indique les classes communes et la façon dont elles sont liées les unes aux autres.

- Un [classeur](/javascript/api/office-scripts/excel/excelscript.workbook) contient une ou plusieurs [feuilles de calcul](/javascript/api/office-scripts/excel/excelscript.worksheet).
- Une [feuille de calcul](/javascript/api/office-scripts/excel/excelscript.worksheet) donne accès à des cellules via [plage](/javascript/api/office-scripts/excel/excelscript.range) objets.
- Une [plage](/javascript/api/office-scripts/excel/excelscript.range) représente un groupe de cellules contiguës.
- Les [plages](/javascript/api/office-scripts/excel/excelscript.range) sont utilisées pour créer et placer des [tableaux](/javascript/api/office-scripts/excel/excelscript.table), des [graphiques](/javascript/api/office-scripts/excel/excelscript.chart), des [formes](/javascript/api/office-scripts/excel/excelscript.shape) et d’autres objets d’organisation ou de visualisation de données.
- Une [feuille de calcul](/javascript/api/office-scripts/excel/excelscript.worksheet) contient des tableaux remplis avec les objets qui sont présents dans la feuille individuelle.
- Un [classeur](/javascript/api/office-scripts/excel/excelscript.workbook) contient des tableaux de certains de ces objets de données pour l’ensemble du classeur.

Pour plus d’informations sur le modèle objet de scripts Office, consultez la [rubrique Scripting Fundamentals for Office scripts in Excel on the Web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>Voir aussi

- [À propos des scripts Office](/office/dev/scripts/overview/excel)
- [Enregistrer, modifier et créer des scripts Office dans Excel sur le web](/office/dev/scripts/tutorials/excel-tutorial)
- [Principes de base des scripts pour Office Scripts dans Excel sur le web](/office/dev/scripts/develop/scripting-fundamentals)
