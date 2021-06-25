# <a name="office-scripts-api-documentation-tools"></a><span data-ttu-id="dde3f-101">Office Outils de documentation de l’API scripts</span><span class="sxs-lookup"><span data-stu-id="dde3f-101">Office Scripts API Documentation Tools</span></span>

<span data-ttu-id="dde3f-102">Ces outils vous aident à Office documentation SCripts et à l’équipe derrière elle.</span><span class="sxs-lookup"><span data-stu-id="dde3f-102">These tools help support the Office SCripts documentation and the team behind it.</span></span> <span data-ttu-id="dde3f-103">Suivez ces instructions pour exécuter les outils dans ce dossier.</span><span class="sxs-lookup"><span data-stu-id="dde3f-103">Follow these instructions to run the tools in this folder.</span></span>

## <a name="coverage-tester"></a><span data-ttu-id="dde3f-104">testeur de couverture</span><span class="sxs-lookup"><span data-stu-id="dde3f-104">coverage-tester</span></span>

<span data-ttu-id="dde3f-105">Cet outil fournit une vue d’ensemble de la couverture de la documentation pour chaque API.</span><span class="sxs-lookup"><span data-stu-id="dde3f-105">This tool gives an overview of the documentation coverage for each API.</span></span> <span data-ttu-id="dde3f-106">Chaque API est évaluée pour la qualité de la documentation et la présence d’exemples de code.</span><span class="sxs-lookup"><span data-stu-id="dde3f-106">Each API is assessed for documentation quality and the presence of sample code.</span></span> <span data-ttu-id="dde3f-107">Les mesures de qualité sont toujours en cours de développement.</span><span class="sxs-lookup"><span data-stu-id="dde3f-107">The quality metrics are still in development.</span></span>

<span data-ttu-id="dde3f-108">La sortie de cet outil est un `.csv` fichier.</span><span class="sxs-lookup"><span data-stu-id="dde3f-108">The output of this tool is a `.csv` file.</span></span>

### <a name="coverage-tester-instructions"></a><span data-ttu-id="dde3f-109">Instructions du testeur de couverture</span><span class="sxs-lookup"><span data-stu-id="dde3f-109">coverage-tester Instructions</span></span>

1. <span data-ttu-id="dde3f-110">Clonez ou bifurcationz le repo.</span><span class="sxs-lookup"><span data-stu-id="dde3f-110">Clone or fork the repo.</span></span>
1. <span data-ttu-id="dde3f-111">Dans une fenêtre de commande, allez à `/office-scripts-docs-reference/generate-docs/tools`</span><span class="sxs-lookup"><span data-stu-id="dde3f-111">In a command window, go to `/office-scripts-docs-reference/generate-docs/tools`</span></span>
1. <span data-ttu-id="dde3f-112">Exécuter `npm install`</span><span class="sxs-lookup"><span data-stu-id="dde3f-112">Run `npm install`</span></span>
1. <span data-ttu-id="dde3f-113">Exécuter `npm run build`</span><span class="sxs-lookup"><span data-stu-id="dde3f-113">Run `npm run build`</span></span>
1. <span data-ttu-id="dde3f-114">Exécuter `node coverage-tester`</span><span class="sxs-lookup"><span data-stu-id="dde3f-114">Run `node coverage-tester`</span></span>
1. <span data-ttu-id="dde3f-115">Ouvrez « Couverture d’API Report.csv »</span><span class="sxs-lookup"><span data-stu-id="dde3f-115">Open “API Coverage Report.csv”</span></span>
