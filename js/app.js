const CLE_REGLES_PLATEAU =
    "planning_gfd_regles_plateau";

let reglesPlateau = [];

let couleursRaf = {};

const CLE_RAF =
    "planning_gfd_raf";
let donneesGlobales = [];

console.log("APP CHARGE");
function afficherOnglet(id) {

    document
        .querySelectorAll(".onglet")
        .forEach(onglet => {
            onglet.classList.add("cache");
        });

    document
    .getElementById(id)
    .classList.remove("cache");

    if(
    id === "plateau"
    ){
    genererPlateauVide();
}
}

document
    .getElementById("dateMaj")
    .textContent =
    "Mise à jour : " +
    new Date().toLocaleString("fr-FR");

async function lireExcel() {

    console.log("DEBUT LECTURE");


    try {

        const response =
            await fetch(
                "data/PLANNING_GFD_UNIQUE_4_TEST.xlsx"
            );

        const arrayBuffer =
            await response.arrayBuffer();

        const workbook =
            XLSX.read(
                arrayBuffer,
                {
                    type: "array"
                }
            );

            chargerReglesPlateau(
                workbook
            );

            chargerCouleursRAF(
                workbook
            );

        console.log(
            "Feuilles trouvées :",
            workbook.SheetNames
        );

            chargerReglesPlateau(
                workbook
        );

        const feuille =
            workbook.Sheets[
                "BDD_GENERALE_PA"
            ];

        const donnees =
            XLSX.utils.sheet_to_json(
                feuille,
                {
                    defval: ""
                }
            );

        donneesGlobales = donnees;

        console.log(
            "Nombre de lignes :",
            donnees.length
        );

        document
            .getElementById(
                "messagePlanning"
            )
            .textContent =
            donnees.length +
            " lignes chargées depuis Excel";

        remplirFiltreRAF(
            donnees
        );

    
        afficherRAF(
            donnees
        );

        afficherReglesPlateau();

        genererPlanningVisuel(
            donnees
        );

    }

catch(erreur){

    console.error(
        "ERREUR :",
        erreur
    );
}
}
function genererPlanningVisuel(donnees) {

    const grille = document.getElementById("grillePlanning");

    const annee = Number(
        document.getElementById("filtreAnnee").value
    );

    const titre =
    document.getElementById(
        "titrePlanning"
    );

    if (titre) {

    titre.textContent =
        "Planning GFD " +
        annee;

}

    const mois = Number(
        document.getElementById("filtreMois").value
    );

    const sectionChoisie =
        document.getElementById("filtreSection").value;
        const filtreRaf =
    document.getElementById(
        "filtreRaf"
    );

    const rafChoisi =
        filtreRaf
        ? filtreRaf.value
        : "TOUS";   

    const premierJour = new Date(
        annee,
        mois,
        1,
        12
    );

    const dernierNumeroJour =
        mois === 11
            ? 30
            : new Date(
                annee,
                mois + 1,
                0
            ).getDate();

    const dernierJour = new Date(
        annee,
        mois,
        dernierNumeroJour,
        12
    );

    const donneesMois =
    donnees.filter(f => {

        const debut =
            convertirDateExcel(
                f["Date de début"]
            );

        const fin =
            convertirDateExcel(
                f["Date de fin"]
            );

        if(
            !debut ||
            !fin
        ){
            return false;
        }

        const sectionOK =

    rafChoisi === "TOUS"

    ||

    normaliserTexte(
        f.RAF
    )

    ===

    normaliserTexte(
        rafChoisi
    )

    ||

    normaliserTexte(
        f.SECTION
    )

    ===

    normaliserTexte(
        sectionChoisie
    );

    const rafOK =

    rafChoisi === "TOUS"

    ||

    normaliserTexte(
        f.RAF
    )

    ===

    normaliserTexte(
        rafChoisi
    );

return (

    sectionOK

    &&

    rafOK

    &&

    debut <= dernierJour

    &&

    fin >= premierJour

);
        

    });

    /*
     * Préparation des formations :
     * conversion des dates Excel et nettoyage des textes.
     */
    const formations = donneesMois
        .map((ligne) => {

            return {
                section:
                    normaliserTexte(
                        ligne.SECTION
                    ),

                domaine:
                    normaliserTexte(
                        ligne.DOMAINES
                    ),

                nom:
                    String(
                        ligne.FORMATIONS || ""
                    ).trim(),

                raf:
                    normaliserTexte(
                        ligne.RAF
                    ),

                debut:
                    convertirDateExcel(
                        ligne["Date de début"]
                    ),

                fin:
                    convertirDateExcel(
                        ligne["Date de fin"]
                    )
            };

        })
        .filter((formation) => {

            if (
                !formation.section ||
                !formation.domaine ||
                !formation.nom ||
                !formation.debut ||
                !formation.fin
            ) {
                return false;
            }

            const sectionValide =
                sectionChoisie === "TOUTES" ||
                formation.section ===
                    normaliserTexte(
                        sectionChoisie
                    );

            const periodeValide =
                formation.debut <= dernierJour &&
                formation.fin >= premierJour;

            return (
                sectionValide &&
                periodeValide
            );

        });

    /*
     * Liste des sections réellement affichées.
     */
    const ordreSections = [
        "DEPARTEMENTALE",
        "NORD",
        "OUEST",
        "SUD"
    ];

    const sections = [
        ...new Set(
            formations.map(
                formation =>
                    formation.section
            )
        )
    ].sort((a, b) => {

        const positionA =
            ordreSections.indexOf(a);

        const positionB =
            ordreSections.indexOf(b);

        if (
            positionA !== -1 ||
            positionB !== -1
        ) {
            if (positionA === -1) {
                return 1;
            }

            if (positionB === -1) {
                return -1;
            }

            return positionA - positionB;
        }

        return a.localeCompare(
            b,
            "fr"
        );

    });

    /*
     * Création des domaines pour chaque section.
     */
    const structure = [];

    sections.forEach((section) => {

        const formationsSection =
            formations.filter(
                formation =>
                    formation.section ===
                    section
            );

        const domaines = [
            ...new Set(
                formationsSection.map(
                    formation =>
                        formation.domaine
                )
            )
        ].sort((a, b) =>
            a.localeCompare(
                b,
                "fr"
            )
        );

        domaines.forEach((domaine) => {

            const formationsDomaine =
                formationsSection
                    .filter(
                        formation =>
                            formation.domaine ===
                            domaine
                    )
                    .sort(
                        (a, b) =>
                            a.debut - b.debut
                    );

            /*
             * Calcul des voies.
             * Deux formations simultanées dans un même domaine
             * sont placées côte à côte.
             */
            const finsDesVoies = [];

            formationsDomaine.forEach(
                (formation) => {

                    const debutVisible =
                        dateMaximum(
                            formation.debut,
                            premierJour
                        );

                    const finVisible =
                        dateMinimum(
                            formation.fin,
                            dernierJour
                        );

                    let voie =
                        finsDesVoies.findIndex(
                            finOccupee =>
                                finOccupee <
                                debutVisible
                        );

                    if (voie === -1) {
                        voie =
                            finsDesVoies.length;
                    }

                    finsDesVoies[voie] =
                        finVisible;

                    formation.voie = voie;
                    formation.debutVisible =
                        debutVisible;
                    formation.finVisible =
                        finVisible;

                }
            );

            structure.push({
                section,
                domaine,
                formations:
                    formationsDomaine,
                nombreVoies:
                    Math.max(
                        1,
                        finsDesVoies.length
                    )
            });

        });

    });

    grille.innerHTML = "";

    if (structure.length === 0) {

        grille.innerHTML = `
            <div class="aucune-formation">
                Aucune formation pour cette sélection.
            </div>
        `;

        return;
    }

    /*
     * Construction des colonnes.
     */
    const LARGEUR_DATE = 82;
    const LARGEUR_VOIE = 145;

    const colonnes = [
        `${LARGEUR_DATE}px`
    ];

    let colonneCourante = 2;

    structure.forEach((element) => {

        element.colonneDebut =
            colonneCourante;

        for (
            let voie = 0;
            voie < element.nombreVoies;
            voie++
        ) {
            colonnes.push(
                `${LARGEUR_VOIE}px`
            );
        }

        colonneCourante +=
            element.nombreVoies;

        element.colonneFin =
            colonneCourante;

    });

    grille.className =
        "grille-planning-visuel";

    grille.style.gridTemplateColumns =
        colonnes.join(" ");

    grille.style.gridTemplateRows =
        `32px repeat(${dernierNumeroJour}, 36px)`;


    /*
     * Bandeaux des sections.
     */
    sections.forEach((section) => {

        const elementsSection =
            structure.filter(
                element =>
                    element.section ===
                    section
            );

        const colonneDebut =
            Math.min(
                ...elementsSection.map(
                    element =>
                        element.colonneDebut
                )
            );

        const colonneFin =
            Math.max(
                ...elementsSection.map(
                    element =>
                        element.colonneFin
                )
            );

        creerCellulePlanning(
            grille,
            `bandeau-section-planning ${classeSection(section)}`,
            libelleSection(section),
            1,
            colonneDebut,
            colonneFin
        );

    });

    /*
     * Deuxième ligne : DATES et domaines.
     */
creerCellulePlanning(
    grille,
    "entete-dates-planning",
    "DATES",
    2,
    1,
    2
);

    structure.forEach((element) => {

        creerCellulePlanning(
            grille,
            "entete-domaine-planning",
            element.domaine,
            2,
            element.colonneDebut,
            element.colonneFin
        );

    });

    /*
     * Création des jours et des cellules de fond.
     */
    for (
        let jour = 1;
        jour <= dernierNumeroJour;
        jour++
    ) {

        const date = new Date(
            annee,
            mois,
            jour,
            12
        );

        const weekEnd =
            date.getDay() === 0 ||
            date.getDay() === 6;

        const ligne = jour + 2;

        const celluleDate =
            document.createElement("div");

        celluleDate.className =
            weekEnd
                ? "jour-planning week-end"
                : "jour-planning";

        celluleDate.textContent =
            `${lettreJour(date)} ${String(jour).padStart(2, "0")}`;

        celluleDate.style.gridRow =
            `${ligne} / ${ligne + 1}`;

        celluleDate.style.gridColumn =
            "1 / 2";

        grille.appendChild(
            celluleDate
        );

        for (
            let colonne = 2;
            colonne < colonnes.length + 1;
            colonne++
        ) {

            const cellule =
                document.createElement(
                    "div"
                );

            cellule.className =
                weekEnd
                    ? "cellule-calendrier week-end"
                    : "cellule-calendrier";

            cellule.style.gridRow =
                `${ligne} / ${ligne + 1}`;

            cellule.style.gridColumn =
                `${colonne} / ${colonne + 1}`;

            grille.appendChild(
                cellule
            );

        }

    }

    /*
     * Placement des formations.
     */
    structure.forEach((element) => {

        element.formations.forEach(
            (formation) => {

                const jourDebut =
                    formation
                        .debutVisible
                        .getDate();

                const jourFin =
                    formation
                        .finVisible
                        .getDate();

                const ligneDebut =
                    jourDebut + 2;

                const ligneFin =
                    jourFin + 3;

                const colonne =
                    element.colonneDebut +
                    formation.voie;

                const bloc =
                    document.createElement(
                        "div"
                    );

                bloc.className =
                    "bloc-formation-planning";

                bloc.style.gridRow =
                    `${ligneDebut} / ${ligneFin}`;

                bloc.style.gridColumn =
                    `${colonne} / ${colonne + 1}`;

                bloc.style.backgroundColor =
                    couleurRaf(
                        formation.raf
                    );

                bloc.style.color =
                    couleurTexte(
                        couleurRaf(
                            formation.raf
                        )
                    );

                bloc.innerHTML = `
                    <strong>
                        ${formation.nom}
                    </strong>

                    ${
                        formation.raf
                            ? `<span>${formation.raf}</span>`
                            : ""
                    }
                `;

                bloc.title =
                    `${formation.nom}\n` +
                    `Section : ${libelleSection(formation.section)}\n` +
                    `Domaine : ${formation.domaine}\n` +
                    `RAF : ${formation.raf || "Non renseigné"}\n` +
                    `${formaterDate(formation.debut)} → ${formaterDate(formation.fin)}`;

                grille.appendChild(
                    bloc
                );

            }
        );

    });

}
function convertirDateExcel(valeur) {

    if (!valeur) {
        return null;
    }

    if (valeur instanceof Date) {
        return new Date(
            valeur.getFullYear(),
            valeur.getMonth(),
            valeur.getDate(),
            12
        );
    }

    if (!Number.isNaN(Number(valeur))) {

        const dateExcel =
            XLSX.SSF.parse_date_code(
                Number(valeur)
            );

        if (!dateExcel) {
            return null;
        }

        return new Date(
            dateExcel.y,
            dateExcel.m - 1,
            dateExcel.d,
            12
        );

    }

    const correspondance =
        String(valeur).match(
            /^(\d{1,2})\d{1,2}[\/.-](\d{4})/
        );

    if (correspondance) {

        return new Date(
            Number(correspondance[3]),
            Number(correspondance[2]) - 1,
            Number(correspondance[1]),
            12
        );

    }

    return null;

}


function normaliserTexte(valeur) {

    return String(valeur || "")
        .trim()
        .toUpperCase()
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        );

}


function creerCellulePlanning(
    grille,
    classe,
    texte,
    ligne,
    colonneDebut,
    colonneFin
) {

    const cellule =
        document.createElement("div");

    cellule.className = classe;
    cellule.textContent = texte;

    cellule.style.gridRow =
        `${ligne} / ${ligne + 1}`;

    cellule.style.gridColumn =
        `${colonneDebut} / ${colonneFin}`;

    grille.appendChild(cellule);

}


function dateMaximum(date1, date2) {

    return date1 > date2
        ? date1
        : date2;

}


function dateMinimum(date1, date2) {

    return date1 < date2
        ? date1
        : date2;

}


function nomMois(numero) {

    const mois = [
        "JANVIER",
        "FÉVRIER",
        "MARS",
        "AVRIL",
        "MAI",
        "JUIN",
        "JUILLET",
        "AOÛT",
        "SEPTEMBRE",
        "OCTOBRE",
        "NOVEMBRE",
        "DÉCEMBRE"
    ];

    return mois[numero];

}


function lettreJour(date) {

    const jours = [
        "D",
        "L",
        "M",
        "Me",
        "J",
        "V",
        "S"
    ];

    return jours[
        date.getDay()
    ];

}


function libelleSection(section) {

    const libelles = {
        DEPARTEMENTALE:
            "DÉPARTEMENTALE",
        NORD:
            "NORD",
        OUEST:
            "OUEST",
        SUD:
            "SUD"
    };

    return libelles[section]
        || section;

}


function classeSection(section) {

    const classes = {
        DEPARTEMENTALE:
            "section-departementale",
        NORD:
            "section-nord",
        OUEST:
            "section-ouest",
        SUD:
            "section-sud"
    };

    return classes[section]
        || "section-autre";

}


function couleurRaf(
    raf
){

    const couleur =
        couleursRaf[
            normaliserTexte(
                raf
            )
        ];

    return couleur
        ||
        "#D9D9D9";

}


function couleurTexte(couleur) {

    const hex =
        couleur.replace("#", "");

    const rouge =
        parseInt(
            hex.substring(0, 2),
            16
        );

    const vert =
        parseInt(
            hex.substring(2, 4),
            16
        );

    const bleu =
        parseInt(
            hex.substring(4, 6),
            16
        );

    const luminosite =
        (
            rouge * 299 +
            vert * 587 +
            bleu * 114
        ) / 1000;

    return luminosite > 145
        ? "#111111"
        : "#FFFFFF";

}


function formaterDate(date) {

    return date.toLocaleDateString(
        "fr-FR"
    );

}
function formaterDateExcel(valeur) {

    if (!valeur) return "";

    const date =
        XLSX.SSF.parse_date_code(
            Number(valeur)
        );

    if (!date) return valeur;

    const jour =
        String(date.d)
        .padStart(2, "0");

    const mois =
        String(date.m)
        .padStart(2, "0");

    const annee =
        date.y;

    return `${jour}/${mois}/${annee}`;

}
function calculerDuree(
    debut,
    fin
){

    const d1 =
        XLSX.SSF.parse_date_code(
            Number(debut)
        );

    const d2 =
        XLSX.SSF.parse_date_code(
            Number(fin)
        );

    if(
        !d1 ||
        !d2
    ){
        return 1;
    }

    const date1 =
        new Date(
            d1.y,
            d1.m - 1,
            d1.d
        );

    const date2 =
        new Date(
            d2.y,
            d2.m - 1,
            d2.d
        );

    const duree =
        Math.round(
            (
                date2 -
                date1
            ) /
            86400000
        ) + 1;

    return Math.max(
        1,
        duree
    );

}

     document
    .getElementById("filtreAnnee")
    .addEventListener(
        "change",
        lireExcel
    );


        document
    .getElementById("filtreMois")
    .addEventListener(
        "change",
        lireExcel
    );



        document
    .getElementById("filtreSection")
    .addEventListener(
        "change",
        lireExcel
    );

        document

    function remplirFiltreRAF(donnees){

    const filtre =
        document.getElementById(
            "filtreRaf"
        );

    if(!filtre){
        return;
    }

    const valeurActuelle =
        filtre.value;

    const rafsExcel =

    donnees
        .map(x => x.RAF)
        .filter(x => x);

    const rafsLocaux =

        Object.keys(
            couleursRaf
        );

    const rafs = [

        ...new Set(

        [
            ...rafsExcel,
            ...rafsLocaux
        ]

    )

].sort();

    console.log(
    "RAF EXCEL :",
    rafsExcel
);

console.log(
    "RAF LOCAUX :",
    rafsLocaux
);

console.log(
    "RAF FUSIONNES :",
    rafs
);

    filtre.innerHTML = `
        <option value="TOUS">
            Tous
        </option>
    `;

    rafs.forEach(raf => {

        filtre.innerHTML += `
            <option value="${raf}">
                ${raf}
            </option>
        `;

    });

    if(
        rafs.includes(
            valeurActuelle
        )
    ){

        filtre.value =
            valeurActuelle;

    }

}

        document
    .getElementById(
        "filtreRaf"
    )
    .addEventListener(
        "change",
        lireExcel
    );

    function chargerCouleursRAF(
    workbook
){

    const feuilleRAF =
        workbook.Sheets["RAF"];

    if(!feuilleRAF){
        return;
    }

    const rafs =
        XLSX.utils.sheet_to_json(
            feuilleRAF,
            {
                defval:""
            }
        );

        rafs.forEach(raf => {

    if(
        raf.RAF &&
        raf.COULEUR
    ){

        const cle =
            normaliserTexte(
                raf.RAF
            );

        if(
            !couleursRaf[cle]
        ){

            couleursRaf[
                cle
            ] =
            raf.COULEUR;

        }

    }

});

}

function afficherRAF(donnees){

    const liste =
        document.getElementById(
            "listeRAF"
        );

    if(!liste){
        return;
    }

    const rafsExcel =

    donnees
        .map(x => x.RAF)
        .filter(x => x);

const rafsLocaux =

    Object.keys(
        couleursRaf
    );

const rafs =

    [
        ...new Set(
            [
                ...rafsExcel,
                ...rafsLocaux
            ]
        )
    ]

    .filter(x => x)

    .sort();

    console.log(
    "RAF EXCEL :",
    rafsExcel
);

console.log(
    "RAF LOCAUX :",
    rafsLocaux
);

console.log(
    "RAF FUSIONNES :",
    rafs
);

    console.log(
    "LISTE RAF AFFICHAGE :",
    rafs
);


    let html = `

    <button
onclick="ajouterRAF()">

    + Ajouter RAF

</button>

<br><br>

    <table class="table-raf">

        <tr>

            <th>RAF</th>

            <th>Aperçu</th>

            <th>Couleur</th>

            <th>Actions</th>

        </tr>

    `;

    rafs.forEach(raf => {

        const couleur =
            couleurRaf(raf);

        html += `

        <tr>

            <td>
                ${raf}
            </td>

            <td>

                <div
                    class="apercu-raf"
                    style="
                    background:${couleur}
                    ">
                </div>

            </td>

            <td>

                <input
                    type="color"
                    value="${couleur}"
                    onchange="
                        changerCouleurRAF(
                            '${raf}',
                            this.value
                        )
                    "
                >

            </td>

            <td>

                 <button
                     onclick="
                        supprimerRAF(
                            '${raf}'
                     )
                     "
                    >
                        Supprimer
                 </button>

            </td>

        </tr>

        `;


    });

    html += "</table>";

    liste.innerHTML = html;

}

function changerCouleurRAF(
    raf,
    couleur
){

    couleursRaf[
        normaliserTexte(raf)
    ] = couleur;

    sauverRAFLocal();

    genererPlanningVisuel(
        donneesGlobales
    );

    afficherRAF(
        donneesGlobales
    );

}

function chargerRAFLocal(){

    const sauvegarde =
        localStorage.getItem(
            CLE_RAF
        );

    if(!sauvegarde){
        return;
    }

    couleursRaf =
        JSON.parse(
            sauvegarde
        );

}

function sauverRAFLocal(){

    localStorage.setItem(
        CLE_RAF,

        JSON.stringify(
            couleursRaf
        )

    );

}

function ajouterRAF(){

    const nom =
        prompt(
            "Nom du RAF"
        );

    if(!nom){
        return;
    }

    couleursRaf[
        normaliserTexte(nom)
    ] =
    "#808080";

    console.log(
    "RAF AJOUTE :",
    couleursRaf
);

    sauverRAFLocal();

afficherRAF(
    donneesGlobales
);

remplirFiltreRAF(
    donneesGlobales
);

}

function supprimerRAF(
    raf
){

    if(
        !confirm(
            "Supprimer " + raf
        )
    ){
        return;
    }

    delete couleursRaf[
        normaliserTexte(raf)
    ];

    sauverRAFLocal();

afficherRAF(
    donneesGlobales
);

remplirFiltreRAF(
    donneesGlobales
);

}

function genererPlateauVide(){

    const annee = Number(
    document.getElementById(
        "filtreAnnee"
    ).value
);

    const zone =
        document.getElementById(
            "planningPlateau"
        );

    if(!zone){
        return;
    }

    zone.innerHTML = "";

    const mois = [

        "JANVIER",
        "FEVRIER",
        "MARS",
        "AVRIL",

        "MAI",
        "JUIN",
        "JUILLET",
        "AOUT",

        "SEPTEMBRE",
        "OCTOBRE",
        "NOVEMBRE",
        "DECEMBRE"

    ];

    const joursCourts = [

        "D",
        "L",
        "M",
        "Me",
        "J",
        "V",
        "S"

    ];

    mois.forEach(
        (nomMois,indexMois)=>{

        let html = `

        <div class="mois-plateau">

            <div class="titre-mois">
                ${nomMois}
            </div>

            <div class="grille-mois">

                <div
                class="entete-colonne">
                    DATES
                </div>

                <div
                class="entete-colonne">
                    INC
                </div>

        `;

        const nbJours =
            new Date(
            annee,
            indexMois + 1,
            0
        ).getDate();

        for(
            let jour=1;
            jour<=nbJours;
            jour++
        ){

        const date =
            new Date(
                annee,
                indexMois,
                jour
            );

            const weekend =

                date.getDay()===0

                ||

                date.getDay()===6;

    html += `

    <div
    class="
    case-date

    ${weekend
        ? "weekend"
        : ""
    }
    ">

        ${jour}

        ${joursCourts[
            date.getDay()
        ]}

    </div>

    <div
    class="case-inc"
    id="m${indexMois}j${jour}">
    </div>

`;

}

        html += `
            </div>
        </div>
        `;

        zone.innerHTML += html;

    });


    afficherPremieresOccupations();

    console.log(
    reglesPlateau
    );
}
    document
    .getElementById(
        "filtreAnnee"
    )
    .addEventListener(
        "change",
        genererPlateauVide
    );

function chargerReglesPlateau(
    workbook
){

    const feuille =
        workbook.Sheets[
            "REGLES PLATEAU"
        ];

    if(!feuille){
        return;
    }

    const reglesExcel =
        XLSX.utils.sheet_to_json(
            feuille,
            {
                defval:""
            }
        );

    /*
     * Si aucune règle n'existe
     * encore dans le navigateur,
     * on importe Excel une seule fois.
     */
    if(
        reglesPlateau.length === 0
    ){

        reglesPlateau =
            reglesExcel;

        sauverReglesPlateau();

        console.log(
            "Import initial des règles depuis Excel"
        );

    }

}

function afficherPremieresOccupations(){

    const filtreAnnee =
        document.getElementById("filtreAnnee");

    const anneeSelectionnee =
        filtreAnnee
            ? Number(filtreAnnee.value)
            : new Date().getFullYear();

    /*
     * BDD_GENERALE_PA peut contenir plusieurs lignes pour une même session.
     * On déduplique donc les sessions avant d'appliquer les règles Plateau.
     */
    const sessionsUniques = new Map();

    donneesGlobales.forEach(formation => {

        const nom =
            normaliserTexte(
                formation.FORMATIONS
            );

        const dateDebut =
            convertirDateExcel(
                formation["Date de début"]
            );

        const dateFin =
            convertirDateExcel(
                formation["Date de fin"]
            );

        if(!nom || !dateDebut){
            return;
        }

        const cleSession = [
            nom,
            dateDebut.getTime(),
            dateFin ? dateFin.getTime() : "",
            normaliserTexte(formation.RAF)
        ].join("|");

        if(!sessionsUniques.has(cleSession)){
            sessionsUniques.set(
                cleSession,
                formation
            );
        }

    });

    sessionsUniques.forEach(formation => {

        const nomFormation =
            normaliserTexte(
                formation.FORMATIONS
            );

        const regle =
            reglesPlateau.find(regleCandidate => {

                const plateauRequis =
                    normaliserTexte(
                        regleCandidate.PLATEAU_REQUIS || "OUI"
                    );

                if(plateauRequis !== "OUI"){
                    return false;
                }

                const nomRegle =
                    normaliserTexte(
                        regleCandidate.FORMATION
                    );

                if(!nomRegle){
                    return false;
                }


                const mode =
                    normaliserTexte(
                        regleCandidate.MODE_CORRESPONDANCE || "CONTIENT"
                    )
                    .replaceAll(" ", "_");

                if(mode === "COMMENCE_PAR"){
                    return nomFormation.startsWith(
                        nomRegle
                    );
                }

                if(
                    mode === "EGAL"
                    ||
                    mode === "EXACT"
                ){
                    return nomFormation === nomRegle;
                }

                if(
    mode ===
    "CONTIENT_TOUS"
){

    const morceaux =
        nomRegle
            .split(";")
            .map(
                x => x.trim()
            );

    return morceaux.every(
        morceau =>

        nomFormation.includes(
            morceau
        )

    );

}

                return nomFormation.includes(
                    nomRegle
                );

            });

        if(!regle){
            return;
        }

        const dateDebut =
            convertirDateExcel(
                formation["Date de début"]
            );

        if(!dateDebut){
            return;
        }

        const joursPlateau = [
            ...new Set(
                String(
                    regle.JOURS_PLATEAU || ""
                )
                .toUpperCase()
                .split(/[;,]/)
                .map(valeur =>
                    Number(
                        valeur
                            .replace("J", "")
                            .trim()
                    )
                )
                .filter(numeroJour =>
                    Number.isInteger(numeroJour)
                    &&
                    numeroJour > 0
                )
            )
        ];

        joursPlateau.forEach(numeroJour => {

            /* J1 = date de début, J2 = début + 1 jour, etc. */
            const dateOccupation =
                new Date(dateDebut);

            dateOccupation.setDate(
                dateOccupation.getDate()
                +
                numeroJour
                -
                1
            );

            /* Le plateau est annuel : on filtre uniquement sur l'année. */
            if(
                dateOccupation.getFullYear()
                !==
                anneeSelectionnee
            ){
                return;
            }

            const mois =
                dateOccupation.getMonth();

            const jour =
                dateOccupation.getDate();

            const cellule =
                document.getElementById(
                    `m${mois}j${jour}`
                );

            if(!cellule){
                return;
            }

            const identifiantBloc = [
                nomFormation,
                dateDebut.getTime(),
                numeroJour
            ].join("|");

            /* Empêche un doublon si la fonction est rappelée sans recréer la grille. */
            const existeDeja =
                Array.from(
                    cellule.querySelectorAll(
                        ".occupation-plateau"
                    )
                )
                .some(bloc =>
                    bloc.dataset.occupationId
                    ===
                    identifiantBloc
                );

            if(existeDeja){
                return;
            }

            const bloc =
                document.createElement("div");

            bloc.className =
                "occupation-plateau";

            bloc.dataset.occupationId =
                identifiantBloc;

            bloc.style.backgroundColor =
                couleurPlateau(
                    formation.FORMATIONS
                );

            bloc.style.borderLeft =
                "4px solid #4472C4";

            bloc.style.padding =
                "2px 4px";

            bloc.style.margin =
                "1px";

            bloc.style.borderRadius =
                "2px";

            bloc.style.fontSize =
                "9px";

            bloc.style.lineHeight =
                "1.1";

            bloc.style.overflow =
                "hidden";

            bloc.title =
                `${formation.FORMATIONS}\n`
                +
                `RAF : ${formation.RAF || "Non renseigné"}\n`
                +
                `Règle : J${numeroJour}\n`
                +
                `Occupation : ${dateOccupation.toLocaleDateString("fr-FR")}`;

            const titre =
    document.createElement("strong");

titre.textContent =
    `${formation.FORMATIONS} - ${formation.RAF || ""}`;

bloc.appendChild(titre);

            cellule.appendChild(bloc);

        });

    });

}

function couleurPlateau(
    formation
){

    const nom =
        formation.toUpperCase();

    if(
        nom.includes("FIP")
    ){
        return "#F4B183";
    }

    if(
        nom.includes("FMPA")
    ){
        return "#B4A7D6";
    }

    if(
        nom.includes("JCI")
    ){
        return "#E06666";
    }

    if(
        nom.includes("SPP")
    ){
        return "#9DC3E6";
    }

    if(
        nom.includes("SPV")
    ){
        return "#A9D18E";
    }

    return "#D9EAD3";

}

function sauverReglesPlateau(){

    localStorage.setItem(

        CLE_REGLES_PLATEAU,

        JSON.stringify(
            reglesPlateau
        )

    );

}

function chargerReglesPlateauLocal(){

    const sauvegarde =

        localStorage.getItem(
            CLE_REGLES_PLATEAU
        );

    if(!sauvegarde){
        return;
    }

    reglesPlateau =
        JSON.parse(
            sauvegarde
        );
    console.log(
    "REGLES LOCALSTORAGE :",
    reglesPlateau.length
);

}

function afficherReglesPlateau(){

    const zone =
        document.getElementById(
            "tableReglesPlateau"
        );

    if(!zone){
        return;
    }

    let html = `

    <table class="table-regles">

        <tr>

            <th>Formation</th>

            <th>Mode</th>

            <th>Plateau</th>

            <th>Jours</th>

            <th>Modifier</th>

            <th>Supprimer</th>

        </tr>

    `;

    reglesPlateau.forEach(
        (regle,index)=>{

        html += `

        <tr>

            <td>
                ${regle.FORMATION}
            </td>

            <td>
                ${regle.MODE_CORRESPONDANCE}
            </td>

            <td>
                ${regle.PLATEAU_REQUIS}
            </td>

            <td>
                ${regle.JOURS_PLATEAU}
            </td>

            <td>

    <button
    onclick="
    modifierReglePlateau(
        ${index}
    )
    ">
        Modifier
    </button>

</td>

<td>

    <button
    onclick="
    supprimerReglePlateau(
        ${index}
    )
    ">
        Supprimer
    </button>

</td>

        </tr>

        `;

    });

    html += `
        </table>
    `;

    zone.innerHTML = html;

}

function ajouterReglePlateau(){

    const formation =
        prompt(
            "Formation"
        );

    if(!formation){
        return;
    }

    const jours =
        prompt(
            "Jours plateau (ex : J2;J3)"
        );

    reglesPlateau.push({

        FORMATION:
            formation,

        MODE_CORRESPONDANCE:
            "COMMENCE_PAR",

        PLATEAU_REQUIS:
            "OUI",

        JOURS_PLATEAU:
            jours

    });

    sauverReglesPlateau();

    afficherReglesPlateau();

}

function modifierReglePlateau(
    index
){

    const regle =
        reglesPlateau[index];

    const formation =
        prompt(
            "Formation",
            regle.FORMATION
        );

    if(!formation){
        return;
    }

    const mode =
        prompt(
            "Mode (COMMENCE_PAR, EGAL, CONTIENT)",
            regle.MODE_CORRESPONDANCE
        );

    if(!mode){
        return;
    }

    const jours =
        prompt(
            "Jours plateau",
            regle.JOURS_PLATEAU
        );

    reglesPlateau[index] = {

        ...regle,

        FORMATION:
            formation,

        MODE_CORRESPONDANCE:
            mode,

        JOURS_PLATEAU:
            jours

    };

    sauverReglesPlateau();

    afficherReglesPlateau();

}

function supprimerReglePlateau(
    index
){

    reglesPlateau.splice(
        index,
        1
    );

    sauverReglesPlateau();

    afficherReglesPlateau();

}


console.log("FIN DU FICHIER");

chargerRAFLocal();
chargerReglesPlateauLocal();


lireExcel();