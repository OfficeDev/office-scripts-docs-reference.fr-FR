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
# <a name="office-scripts-api-reference"></a><span data-ttu-id="59a9b-103">Référence de l'API Office Scripts</span><span class="sxs-lookup"><span data-stu-id="59a9b-103">Office Scripts API reference</span></span>

<span data-ttu-id="59a9b-104">L’API de scripts Office vous permet d’automatiser des tâches courantes dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="59a9b-104">The Office Scripts API lets you automate common tasks in Excel on the web.</span></span> <span data-ttu-id="59a9b-105">Utilisez cette documentation de référence pour en savoir plus sur les classes, les méthodes et les autres types disponibles pour vos scripts.</span><span class="sxs-lookup"><span data-stu-id="59a9b-105">Use this reference documentation to learn more about the classes, methods, and other types available for your scripts.</span></span> <span data-ttu-id="59a9b-106">Tous les objets accessibles par le biais de scripts Office sont disponibles dans la table des matières sur la gauche de la page.</span><span class="sxs-lookup"><span data-stu-id="59a9b-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

> [!NOTE]
> <span data-ttu-id="59a9b-107">Si vous recherchez les API JavaScript pour le développement de compléments Office, consultez la référence de l' [API JavaScript pour les compléments Office](/javascript/api/overview?view=excel-js-preview).</span><span class="sxs-lookup"><span data-stu-id="59a9b-107">If you're looking for the JavaScript APIs for developing Office Add-ins, visit the [Office Add-ins JavaScript API reference](/javascript/api/overview?view=excel-js-preview).</span></span>

## <a name="common-classes"></a><span data-ttu-id="59a9b-108">Classes communes</span><span class="sxs-lookup"><span data-stu-id="59a9b-108">Common classes</span></span>

<span data-ttu-id="59a9b-109">La liste suivante décompose les concepts de base du modèle objet de scripts Office.</span><span class="sxs-lookup"><span data-stu-id="59a9b-109">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="59a9b-110">Cela indique les classes communes et la façon dont elles sont liées les unes aux autres.</span><span class="sxs-lookup"><span data-stu-id="59a9b-110">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="59a9b-111">Un [classeur](/javascript/api/office-scripts/excel/excelscript.workbook) contient une ou plusieurs [feuilles de calcul](/javascript/api/office-scripts/excel/excelscript.worksheet).</span><span class="sxs-lookup"><span data-stu-id="59a9b-111">A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excelscript.worksheet).</span></span>
- <span data-ttu-id="59a9b-112">Une [feuille de calcul](/javascript/api/office-scripts/excel/excelscript.worksheet) donne accès à des cellules via [plage](/javascript/api/office-scripts/excel/excelscript.range) objets.</span><span class="sxs-lookup"><span data-stu-id="59a9b-112">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excelscript.range) objects.</span></span>
- <span data-ttu-id="59a9b-113">Une [plage](/javascript/api/office-scripts/excel/excelscript.range) représente un groupe de cellules contiguës.</span><span class="sxs-lookup"><span data-stu-id="59a9b-113">A [Range](/javascript/api/office-scripts/excel/excelscript.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="59a9b-114">Les [plages](/javascript/api/office-scripts/excel/excelscript.range) sont utilisées pour créer et placer des [tableaux](/javascript/api/office-scripts/excel/excelscript.table), des [graphiques](/javascript/api/office-scripts/excel/excelscript.chart), des [formes](/javascript/api/office-scripts/excel/excelscript.shape) et d’autres objets d’organisation ou de visualisation de données.</span><span class="sxs-lookup"><span data-stu-id="59a9b-114">[Ranges](/javascript/api/office-scripts/excel/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excelscript.table), [Charts](/javascript/api/office-scripts/excel/excelscript.chart), [Shapes](/javascript/api/office-scripts/excel/excelscript.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="59a9b-115">Une [feuille de calcul](/javascript/api/office-scripts/excel/excelscript.worksheet) contient des tableaux remplis avec les objets qui sont présents dans la feuille individuelle.</span><span class="sxs-lookup"><span data-stu-id="59a9b-115">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) contains arrays filled with those objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="59a9b-116">Un [classeur](/javascript/api/office-scripts/excel/excelscript.workbook) contient des tableaux de certains de ces objets de données pour l’ensemble du classeur.</span><span class="sxs-lookup"><span data-stu-id="59a9b-116">A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains arrays of some of those data objects for the entire Workbook.</span></span>

<span data-ttu-id="59a9b-117">Pour plus d’informations sur le modèle objet de scripts Office, consultez la [rubrique Scripting Fundamentals for Office scripts in Excel on the Web](/office/dev/scripts/develop/scripting-fundamentals)</span><span class="sxs-lookup"><span data-stu-id="59a9b-117">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="59a9b-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="59a9b-118">See also</span></span>

- [<span data-ttu-id="59a9b-119">À propos des scripts Office</span><span class="sxs-lookup"><span data-stu-id="59a9b-119">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="59a9b-120">Enregistrer, modifier et créer des scripts Office dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="59a9b-120">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="59a9b-121">Principes de base des scripts pour Office Scripts dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="59a9b-121">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
