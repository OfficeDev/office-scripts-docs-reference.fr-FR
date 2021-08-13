---
title: Référence de l'API Office Scripts
description: Vue d’ensemble des API JavaScript Office scripts.
ms.date: 06/29/2020
ms.openlocfilehash: 3ce3344fb49b2811719feb13f8fb4118f1a20060db9d85a06d1be939f22bf3c5
ms.sourcegitcommit: 6a0182075a558c4fe664fedfee08fea76513b192
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2021
ms.locfileid: "58183425"
---
# <a name="office-scripts-api-reference"></a>Référence de l'API Office Scripts

L Office API Scripts vous permet d’automatiser les tâches courantes dans Excel sur le Web. Utilisez cette documentation de référence pour en savoir plus sur les classes, méthodes et autres types disponibles pour vos scripts. Tous les objets accessibles via Office Scripts sont accessibles dans la table des matières à gauche de la page.

> [!NOTE]
> Si vous recherchez les API JavaScript pour le développement de Office, visitez la Office de [l’API JavaScript](/javascript/api/overview?view=excel-js-preview)pour les Office.

## <a name="common-classes"></a>Classes courantes

La liste suivante décompose les principes de base du modèle objet Office Scripts. Cela montre les classes communes et leur relation les unes avec les autres.

- Un [classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) contient une ou plusieurs [feuilles de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet).
- Une [feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet) donne accès à des cellules via [plage](/javascript/api/office-scripts/excelscript/excelscript.range) objets.
- Une [plage](/javascript/api/office-scripts/excelscript/excelscript.range) représente un groupe de cellules contiguës.
- Les [plages](/javascript/api/office-scripts/excelscript/excelscript.range) sont utilisées pour créer et placer des [tableaux](/javascript/api/office-scripts/excelscript/excelscript.table), des [graphiques](/javascript/api/office-scripts/excelscript/excelscript.chart), des [formes](/javascript/api/office-scripts/excelscript/excelscript.shape) et d’autres objets d’organisation ou de visualisation de données.
- Une [feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contient des tableaux contenant les objets présents dans la feuille individuelle.
- Un [workbook contient](/javascript/api/office-scripts/excelscript/excelscript.workbook) des tableaux de certains de ces objets de données pour l’ensemble du workbook.

Pour plus d’informations sur le modèle objet Office Scripts, consultez les principes de base des [scripts Office scripts dans Excel sur le Web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>Voir aussi

- [À propos Office scripts](/office/dev/scripts/overview/excel)
- [Enregistrer, modifier et créer des scripts Office dans Excel sur le web](/office/dev/scripts/tutorials/excel-tutorial)
- [Principes de base pour la rédaction de scripts Office en Excel sur le web](/office/dev/scripts/develop/scripting-fundamentals)
