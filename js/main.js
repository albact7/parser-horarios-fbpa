var file = document.getElementById('docpicker')
var viewer = document.getElementById('dataviewer')
file.addEventListener('change', importFile);

function importFile(evt) {
    var f = evt.target.files[0];

    if (f) {
        var r = new FileReader();
        r.onload = e => {
            var contents = processExcel(e.target.result);
            var partidos = []
            if (contents['Hoja1']) {
                contents['Hoja1'].filter(function(el) {
                    partido = []
                    el.filter(function(el2) {
                        if (el2.toString().normalize() != "") partido.push(el2)

                    });
                    if (partido.length > 0) partidos.push(partido)
                });
                writeInHTML(partidos)
            } else {
                var workbook = XLSX.read(e.target.result, {
                    type: 'binary'
                });
                var strings = workbook.Strings;
                var i = 0
                for (i = 8; i < strings.length; i += 6) {
                    console.log(i + " " + strings[i + 5].t)
                    $('#horariospartidos').append('<div class="categoria">' +
                        '<div class="titulocategoria">' +
                        '<p class="titulocategoria">' +
                        strings[i].t + '</p>' +
                        '</div>' +
                        '<p id="partido" class="p-partido local">' + strings[i + 1].t + '</p>' +
                        '<p id="partido" class="p-partido fecha">' + strings[i + 2].t + '</p>' +
                        '<p id="partido" class="p-partido lugar">' + strings[i + 3].t + '</p>' +
                        '<p id="partido" class="p-partido visitante">' + strings[i + 5].t + '</p>' +
                        '</div>');
                }
            }
        }
        r.readAsBinaryString(f);
    } else {
        console.log("Failed to load file");
    }
}

function processExcel(data) {
    var workbook = XLSX.read(data, {
        type: 'binary'
    });
    var data = to_json(workbook);
    console.log(workbook)
    return data
};

function to_json(workbook) {
    var result = {};
    workbook.SheetNames.forEach(function(sheetName) {
        var roa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
            header: 1
        });
        if (roa.length) result[sheetName] = roa;
    });
    return result; //JSON.stringify(result, 2, 2);
};

function writeInHTML(partidos) {
    var i = 0;
    var back = 0;
    var info_partido = "info_partidoA";
    for (i = 4; i < partidos.length; i += 2) {
        if (back % 2 == 0) {
            info_partido = "info_partidoA";
        } else {
            info_partido = "info_partidoB";
        }
        back++;
        $('#horariospartidos').append('<div class="categoria ' + info_partido + '">' +
            '<div class="titulocategoria">' +
            '<p class="titulocategoria">' +
            partidos[i][0] + '</p>' +
            '</div>' +
            '<div class="row"><p id="partido" class="p-partido local col-3">' + partidos[i][1] + '</p>' +
            '<p id="partido" class="p-partido visitante col-3">' + partidos[i + 1][0] + '</p>' +
            '<p id="partido" class="p-partido fecha col-2">' + partidos[i][2] + '</p>' +
            '<p id="partido" class="p-partido fecha col-2">' + partidos[i + 1][1] + '</p>' +
            '<p id="partido" class="p-partido pabellon col-2">' + partidos[i][3] + '</p></div>' +

            '</div>');
    }

}

$('#printJPG').click(function() {

    var w = document.getElementById("horariospartidos").offsetWidth;
    var h = document.getElementById("horariospartidos").offsetHeight;
    html2canvas(document.getElementById("horariospartidos"), {
        allowTaint: true,
        foreignObjectRendering: true,
        imageSmoothingEnabled: false,
        mozImageSmoothingEnabled: false,
        oImageSmoothingEnabled: false,
        webkitImageSmoothingEnabled: false,
        msImageSmoothingEnabled: false,
        dpi: 300,
        scale: 2,
        onrendered: function(canvas) {
            var a = document.createElement('a');
            a.href = canvas.toDataURL("image/jpeg", 1).replace("image/jpeg", "image/octet-stream");
            a.download = 'horarios.jpg';
            a.click();
        }
    });
});

$('#printPDF').click(function() {

    var w = document.getElementById("horariospartidos").offsetWidth;
    var h = document.getElementById("horariospartidos").offsetHeight;
    var HTML_Width = $("#horariospartidos").width() * 3;
    var HTML_Height = $("#horariospartidos").height() * 3;
    var top_left_margin = 15;
    if (HTML_Width >= HTML_Height) {
        var PDF_Width = HTML_Width + (top_left_margin * 2);
        var PDF_Height = (PDF_Width) + (top_left_margin * 2);
    } else {
        var PDF_Width = HTML_Width + (top_left_margin * 2);
        var PDF_Height = (HTML_Height) + (top_left_margin * 2);
    }
    var canvas_image_width = HTML_Width;
    var canvas_image_height = HTML_Height;
    html2canvas(document.getElementById("horariospartidos"), {
        orientation: "landscape",
        allowTaint: true,
        foreignObjectRendering: true,
        dpi: 300,
        scale: 1,
        quality: 4,
        onrendered: function(canvas) {
            var imgData = canvas.toDataURL('image/png', 1.0);
            var doc = new jsPDF('p', 'pt', [PDF_Width, PDF_Height]);
            doc.addImage(imgData, 'PNG', top_left_margin, top_left_margin, canvas_image_width, canvas_image_height);
            doc.save('horarios.pdf');
        }
    });
});