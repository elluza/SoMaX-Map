Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
        // Excel on valmis, haetaan solun arvo
        Excel.run(function (context) {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange("A1"); // Esimerkki: haetaan solun A1 arvo
            range.load("values");  // Ladataan solun arvot

            return context.sync().then(function () {
                const cellValue = range.values[0][0]; // Solun A1 arvo
                console.log("Solun A1 arvo on: " + cellValue);

                // Käytetään arvoa, esim. päivitetään HTML-sivu
                document.getElementById("map-container").innerHTML = 
                    "Solun A1 arvo on: " + cellValue;
            });
        }).catch(function (error) {
            console.log("Virhe: " + error);
        });
    }
});

let map;
document.addEventListener('DOMContentLoaded', function () {
    map = L.map('map', {
        center: [62.2416, 25.7209], 
        zoom: 7,
        zoomControl: true,
        scrollWheelZoom: false,
        doubleClickZoom: false,
        dragging: false,
        touchZoom: false,
    });

    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
    }).addTo(map);
});

function initializeMap() {
    map.scrollWheelZoom.enable();
    map.doubleClickZoom.enable();
    map.dragging.enable();
    map.touchZoom.enable();
    map.zoomControl.enable();
    document.getElementById('map').classList.add('interactive');
    alert("Kartta on nyt interaktiivinen!");
}

