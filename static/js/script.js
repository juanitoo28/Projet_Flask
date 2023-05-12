/* 
Ce fichier JavaScript fait partie de l'application 'mon_application' du projet Django 'mon_projet'. 

Il définit plusieurs fonctions qui sont utilisées pour interagir avec l'interface utilisateur, telles que :
1. L'ajout d'une nouvelle feuille à un fichier Excel lorsqu'un bouton est cliqué ('runFunction' event listener).
2. La navigation fluide entre les sections de la page grâce à l'écouteur d'événements de navigation.
3. La lecture d'un fichier Excel côté client et l'affichage de son aperçu.
4. La lecture d'un fichier Excel côté serveur et l'affichage de son aperçu.
5. L'affichage de l'aperçu d'un fichier Excel dans un tableau HTML.
6. L'écouteur d'événements qui déclenche la lecture du fichier Excel et son aperçu lorsque l'utilisateur sélectionne un fichier.

Pour plus d'informations sur les concepts utilisés dans ce fichier, tels que les Fetch API, les promesses, async/await et les événements, consultez :
- Fetch API : https://developer.mozilla.org/fr/docs/Web/API/Fetch_API
- Promesses et async/await : https://developer.mozilla.org/fr/docs/Learn/JavaScript/Asynchronous
- Événements : https://developer.mozilla.org/fr/docs/Web/API/Event
- FileReader API : https://developer.mozilla.org/fr/docs/Web/API/FileReader
*/

document.getElementById("runFunction").addEventListener("click", async function () {
    try {
        const response = await fetch("/add_sheet/", {
            method: "POST",
        });

        if (response.status === 200) {
            alert("Feuille ajoutée avec succès !");
        } else {
            alert("Erreur lors de l'ajout de la feuille.");
        }
    } catch (error) {
        console.error(error);
        alert("Erreur lors de l'ajout de la feuille.");
    }
});


document.querySelectorAll("nav a").forEach((link) => {
    link.addEventListener("click", (event) => {
        event.preventDefault();
        const targetId = event.target.getAttribute("href");
        const targetElement = document.querySelector(targetId);
        if (targetElement) {
            targetElement.scrollIntoView({ behavior: "smooth" });
        }
    });
});

function readExcelFile(file) {
    const fileReader = new FileReader();
    fileReader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        displayExcelPreview(sheetData, "côté client");
    };
    fileReader.readAsArrayBuffer(file);
}

async function readExcelFileServerSide(file) {
    const formData = new FormData();
    formData.append("excel_file", file);

    try {
        const response = await fetch("/excel_preview/", {
            method: "POST",
            body: formData
        });

        if (response.status === 200) {
            const data = await response.json();
            displayExcelPreview(data.data, "côté serveur");
        } else {
            alert("Erreur lors de l'aperçu du fichier.");
        }
    } catch (error) {
        console.error(error);
        alert("Erreur lors de l'aperçu du fichier.");
    }
}


function displayExcelPreview(sheetData, source) {
    const previewContainer = document.getElementById("file_preview");
    const table = document.createElement("table");
    const sourceText = document.createElement("h3");
    sourceText.textContent = `Aperçu du fichier Excel (${source})`;

    sheetData.forEach((rowData) => {
        const row = document.createElement("tr");

        rowData.forEach((cellData) => {
            const cell = document.createElement("td");
            cell.textContent = cellData;
            row.appendChild(cell);
        });

        table.appendChild(row);
    });

    previewContainer.innerHTML = "";
    previewContainer.appendChild(sourceText);
    previewContainer.appendChild(table);
}


document.getElementById("excel_file").addEventListener("change", function (event) {
    const file = event.target.files[0];
    const previewMethod = document.querySelector('input[name="preview_method"]:checked').value;

    if (file) {
        if (previewMethod === "client_side") {
            readExcelFile(file);
        } else {
            readExcelFileServerSide(file);
        }
    } else {
        document.getElementById("file_preview").innerHTML = "";
    }
});
