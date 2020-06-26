---
title: Référence de l'API Office Scripts
description: Vue d’ensemble des API JavaScript Async pour les scripts Office.
ms.date: 06/17/2020
ms.openlocfilehash: b9fa59764b7cb54567adbe05b9671bae9b232972
ms.sourcegitcommit: 163b26a43411ad7f13a01237efe9b8d6de656b47
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44883140"
---
# <a name="office-scripts-async-api-reference"></a>Référence de l’API asynchrone de scripts Office

L’API Async de scripts Office prend en charge les scripts plus anciens réalisés pendant la phase d’aperçu des scripts Office. Utilisez cette documentation de référence pour en savoir plus sur les classes, les méthodes et les autres types utilisés par ces anciens scripts. Tous les objets accessibles par le biais de scripts Office sont disponibles dans la table des matières sur la gauche de la page.

> [!IMPORTANT]
> Nous vous recommandons vivement de créer des scripts avec les API de scripts Office standard. Si vous créez ou modifiez un nouveau script, passez à la [version non-Async](?view=office-scripts) des API.

## <a name="common-classes"></a>Classes communes

La liste suivante décompose les concepts de base du modèle objet de scripts Office. Cela indique les classes communes et la façon dont elles sont liées les unes aux autres.

- Un [classeur](/javascript/api/office-scripts/excel/excelscript.workbook) contient une ou plusieurs [feuilles de calcul](/javascript/api/office-scripts/excel/excelscript.worksheet) dans un [WorksheetCollection](/javascript/api/office-scripts/excel/excelscript.worksheetcollection).
- Une [feuille de calcul](/javascript/api/office-scripts/excel/excelscript.worksheet) donne accès à des cellules via [plage](/javascript/api/office-scripts/excel/excelscript.range) objets.
- Une [plage](/javascript/api/office-scripts/excel/excelscript.range) représente un groupe de cellules contiguës.
- Les [plages](/javascript/api/office-scripts/excel/excelscript.range) sont utilisées pour créer et placer des [tableaux](/javascript/api/office-scripts/excel/excelscript.table), des [graphiques](/javascript/api/office-scripts/excel/excelscript.chart), des [formes](/javascript/api/office-scripts/excel/excelscript.shape) et d’autres objets d’organisation ou de visualisation de données.
- Une [feuille de calcul](/javascript/api/office-scripts/excel/excelscript.worksheet) contient des collections de ces objets de données (tels qu’un [ChartCollection](/javascript/api/office-scripts/excel/excelscript.chartcollection)) qui sont présents dans la feuille individuelle.
- Les [classeurs](/javascript/api/office-scripts/excel/excelscript.workbook) contiennent des collections de certains de ces objets de données (par exemple, un [TableCollection](/javascript/api/office-scripts/excel/excelscript.tablecollection)) pour l’intégralité du [classeur](/javascript/api/office-scripts/excel/excelscript.workbook).

Pour plus d’informations sur le modèle objet de scripts Office, consultez la [rubrique Scripting Fundamentals for Office scripts in Excel on the Web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>Voir aussi

- [Utilisation des API Async de scripts Office pour prendre en charge les scripts hérités](/office/dev/scripts/develop/excel-async-model)
- [À propos des scripts Office](/office/dev/scripts/overview/excel)
- [Enregistrer, modifier et créer des scripts Office dans Excel sur le web](/office/dev/scripts/tutorials/excel-tutorial)
- [Principes de base des scripts pour Office Scripts dans Excel sur le web](/office/dev/scripts/develop/scripting-fundamentals)
