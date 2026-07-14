let grafico = null;
let graficoBarras = null;
let total = 0;
let alto = 0;
let medio = 0;
let bajo = 0;

let promedioAsistencia = 0;
let promedioGeneral = 0;
let promedioTareas = 0;

let indicadorCritico = "";
let listaEstudiantes = [];
let resumenIA = "";
const archivo = document.getElementById("archivoExcel");

archivo.addEventListener("change", leerExcel);

function leerExcel(evento){

    const archivoSeleccionado = evento.target.files[0];

    const lector = new FileReader();

    lector.onload = function(e){

        const datos = new Uint8Array(e.target.result);

        const libro = XLSX.read(datos,{type:"array"});

        const hoja = libro.Sheets[libro.SheetNames[0]];

        const estudiantes = XLSX.utils.sheet_to_json(hoja);

        listaEstudiantes = estudiantes;

        console.log(estudiantes);

        const cuerpoTabla = document.querySelector("#tablaEstudiantes tbody");

        cuerpoTabla.innerHTML = "";

        let total = estudiantes.length;

            let alto = 0;

            let medio = 0;

            let bajo = 0;

            let sumaAsistencia = 0;

            let sumaPromedio = 0;

            let sumaTareas = 0;

        estudiantes.forEach(estudiante => {

            sumaAsistencia += Number(estudiante.Asistencia);

            sumaPromedio += Number(estudiante.Promedio);

            sumaTareas += Number(estudiante.Tareas);

            const fila = document.createElement("tr");

            let riesgo = "BAJO";
            let claseRiesgo = "riesgo-bajo";
            let recomendacion = "";
            let motivo = "";

            if (
                estudiante.Asistencia < 70 ||
                estudiante.Promedio < 11 ||
                estudiante.Tareas < 60
            ){
                riesgo = "ALTO";
                claseRiesgo = "riesgo-alto";

                let recomendaciones = [];

                let motivos = [];

                if(estudiante.Asistencia < 70){

                    motivos.push("Baja asistencia");

                    recomendaciones.push("Reunión con padres");

                }

                if(estudiante.Promedio < 11){

                    motivos.push("Bajo promedio");

                    recomendaciones.push("Tutoría personalizada");

                }

                if(estudiante.Tareas < 60){

                    motivos.push("Pocas tareas entregadas");

                    recomendaciones.push("Plan de recuperación de tareas");

                }

                motivo = motivos.join(", ");

                recomendacion = recomendaciones.join(", ");

                alto++;

            }
            else if(
                estudiante.Asistencia < 85 ||
                estudiante.Promedio < 14 ||
                estudiante.Tareas < 80
            ){
                riesgo = "MEDIO";
                claseRiesgo = "riesgo-medio";

                let recomendaciones = [];

                let motivos = [];

                if(estudiante.Asistencia < 85){

                    motivos.push("Asistencia mejorable");

                    recomendaciones.push("Control semanal de asistencia");

                }

                if(estudiante.Promedio < 14){

                    motivos.push("Promedio regular");

                    recomendaciones.push("Reforzamiento académico");

                }

                if(estudiante.Tareas < 80){

                    motivos.push("Entrega parcial de tareas");

                    recomendaciones.push("Seguimiento de tareas");

                }

                motivo = motivos.join(", ");

                recomendacion = recomendaciones.join(", ");

                medio++;

            }
            else{
                riesgo = "BAJO";
                claseRiesgo = "riesgo-bajo";

                motivo = "Buen desempeño general";

                recomendacion = "Mantener el rendimiento y reconocer el esfuerzo.";

                bajo++;

            }

            fila.innerHTML = `
                <td>${estudiante.Nombre}</td>
                <td>${estudiante.Asistencia}</td>
                <td>${estudiante.Promedio}</td>
                <td>${estudiante.Participación}</td>
                <td>${estudiante.Conducta}</td>
                <td>${estudiante.Tareas}</td>
                <td>
                    <span class="${claseRiesgo}">
                        ${riesgo}
                    </span>
                </td>

                <td>${motivo}</td>

                <td>${recomendacion}</td>
            `;

            cuerpoTabla.appendChild(fila);

                fila.addEventListener("click", function(){

                    const filas = document.querySelectorAll("#tablaEstudiantes tbody tr");

                    filas.forEach(f=>{
                        f.classList.remove("fila-seleccionada");
                    });

                    fila.classList.add("fila-seleccionada");

                    document.getElementById("detalleEstudiante").innerHTML = `

                        <h2>👤 ${estudiante.Nombre}</h2>

                        <h3>Estado Académico</h3>

                        <p><strong>Nivel de riesgo:</strong>
                        <span class="${claseRiesgo}">
                        ${riesgo}
                        </span>
                        </p>

                        <h3>Indicadores</h3>

                        <p><strong>📊 Asistencia:</strong> ${estudiante.Asistencia}%</p>

                        <p><strong>📚 Promedio:</strong> ${estudiante.Promedio}</p>

                        <p><strong>🙋 Participación:</strong> ${estudiante.Participación}</p>

                        <p><strong>😊 Conducta:</strong> ${estudiante.Conducta}</p>

                        <p><strong>📝 Tareas:</strong> ${estudiante.Tareas}%</p>

                        <h3>Motivos Detectados</h3>

                        <p>${motivo}</p>

                        <h3>Recomendaciones</h3>

                        <p>${recomendacion}</p>

                    `;

                });


        });
            document.getElementById("totalEstudiantes").textContent = total;

            document.getElementById("riesgoAlto").textContent = alto;

            document.getElementById("riesgoMedio").textContent = medio;

            document.getElementById("riesgoBajo").textContent = bajo;

            promedioAsistencia = (sumaAsistencia / total).toFixed(1);
            promedioGeneral = (sumaPromedio / total).toFixed(1);
            promedioTareas = (sumaTareas / total).toFixed(1);

            const ctx = document.getElementById("graficoRiesgo");

            if(grafico){

                grafico.destroy();

            }

            grafico = new Chart(ctx,{

                type:"pie",

                data:{

                    labels:["Riesgo Alto","Riesgo Medio","Riesgo Bajo"],

                    datasets:[{

                        data:[alto,medio,bajo],

                        backgroundColor:[

                            "#e53935",
                            "#fbc02d",
                            "#43a047"

                        ]

                    }]

                },

                options:{

                    responsive:true,

                    plugins:{

                        legend:{

                            position:"bottom"

                        },

                        title:{

                            display:true,

                            text:"Distribución del Riesgo Académico"

                        }

                    }

                }

            });


            const contextoBarras = document.getElementById("graficoIndicadores");

            if(graficoBarras){

                graficoBarras.destroy();
                
            }

            graficoBarras = new Chart(contextoBarras,{

                type:"bar",

                data:{

                    labels:[
                        "Asistencia",
                        "Promedio",
                        "Tareas"
                    ],

                    datasets:[{

                        label:"Promedio General",
                        data:[
                            promedioAsistencia,
                            promedioGeneral,
                            promedioTareas
                        ],
                    }]
                },

                options:{

                    responsive:true,
                    plugins:{
                        title:{
                            display:true,
                            text:"Indicadores generales del aula"
                        }
                    }
                }
            });

            let indicadorCritico = "Asistencia";

            let valorCritico = promedioAsistencia;

            if(promedioGeneral < valorCritico){

                indicadorCritico = "Promedio";

                valorCritico = promedioGeneral;

            }

            if(promedioTareas < valorCritico){

                indicadorCritico = "Tareas";

            }

            document.getElementById("resumenAnalisis").innerHTML = `

            <p><strong>Total de estudiantes:</strong> ${total}</p>

            <p><strong>Asistencia promedio:</strong> ${promedioAsistencia}%</p>

            <p><strong>Promedio general:</strong> ${promedioGeneral}</p>

            <p><strong>Tareas promedio:</strong> ${promedioTareas}%</p>

            <p><strong>Riesgo Alto:</strong> ${alto} estudiantes (${((alto/total)*100).toFixed(1)}%)</p>

            <p><strong>Indicador más crítico:</strong> ${indicadorCritico}</p>

            `;

let prioridad = "🟢 BAJA";

if(alto >= total*0.30){

    prioridad = "🔴 ALTA";

}else if(medio >= total*0.30){

    prioridad = "🟡 MEDIA";

}

resumenIA = `
<h3>📊 Diagnóstico automático</h3>

<p>Se analizaron <strong>${total}</strong> estudiantes.</p>

<p>Se detectaron <strong>${alto}</strong> estudiantes con riesgo alto.</p>

<p>El promedio de asistencia es <strong>${promedioAsistencia}%</strong>.</p>

<p>El promedio general es <strong>${promedioGeneral}</strong>.</p>

<p>El promedio de tareas es <strong>${promedioTareas}%</strong>.</p>

<hr>

<h3>🤖 Recomendaciones</h3>
`;

if(promedioAsistencia < 85){

    resumenIA += `
    <p>✔ Mejorar el control de asistencia y realizar seguimiento semanal.</p>
    <p>✔ Coordinar reuniones con padres de familia.</p>
    `;

}

if(promedioGeneral < 14){

    resumenIA += `
    <p>✔ Implementar tutorías académicas.</p>
    <p>✔ Realizar reforzamiento después de clases.</p>
    `;

}

if(promedioTareas < 80){

    resumenIA += `
    <p>✔ Supervisar la entrega de tareas.</p>
    <p>✔ Aplicar estrategias de motivación.</p>
    `;

}

resumenIA += `
<hr>

<h3>Prioridad del aula: ${prioridad}</h3>
`;
            
const ranking = [...estudiantes];

ranking.sort((a, b) => {

    let puntajeA = 0;
    let puntajeB = 0;

    if (a.Asistencia < 70) puntajeA += 3;
    else if (a.Asistencia < 85) puntajeA += 2;

    if (a.Promedio < 11) puntajeA += 3;
    else if (a.Promedio < 14) puntajeA += 2;

    if (a.Tareas < 60) puntajeA += 3;
    else if (a.Tareas < 80) puntajeA += 2;

    if (b.Asistencia < 70) puntajeB += 3;
    else if (b.Asistencia < 85) puntajeB += 2;

    if (b.Promedio < 11) puntajeB += 3;
    else if (b.Promedio < 14) puntajeB += 2;

    if (b.Tareas < 60) puntajeB += 3;
    else if (b.Tareas < 80) puntajeB += 2;

    return puntajeB - puntajeA;

});

let htmlRanking = "<ol>";

ranking.slice(0,5).forEach(estudiante=>{

    htmlRanking += `<li><strong>${estudiante.Nombre}</strong></li>`;

});

htmlRanking += "</ol>";

document.getElementById("rankingEstudiantes").innerHTML = htmlRanking;

    };

    lector.readAsArrayBuffer(archivoSeleccionado);

}

const buscador = document.getElementById("buscar");

buscador.addEventListener("keyup", function(){

    const texto = buscador.value.toLowerCase();

    const filas = document.querySelectorAll("#tablaEstudiantes tbody tr");

    filas.forEach(fila => {

        const nombre = fila.cells[0].textContent.toLowerCase();

        if(nombre.includes(texto)){

            fila.style.display = "";

        }else{

            fila.style.display = "none";

        }

    });

});

const botonTodos = document.getElementById("todos");

const botonAlto = document.getElementById("alto");

const botonMedio = document.getElementById("medio");

const botonBajo = document.getElementById("bajo");

function filtrarPorRiesgo(tipo){

    const filas = document.querySelectorAll("#tablaEstudiantes tbody tr");

    filas.forEach(fila=>{

        const riesgo = fila.cells[6].textContent.trim().toUpperCase();

        if(tipo==="TODOS"){

            fila.style.display="";

        }else if(riesgo===tipo){

            fila.style.display="";

        }else{

            fila.style.display="none";

        }

    });

}

const botonIA = document.getElementById("analizarIA");

const resultadoIA = document.getElementById("resultadoIA");
botonIA.addEventListener("click",function(){

    resultadoIA.innerHTML = resumenIA;

});
function generarAnalisisIA(){

    let mensaje = "";

    let prioridad = "🟢 BAJA";

    if(alto > total*0.30){

        prioridad = "🔴 ALTA";

    }else if(medio > total*0.30){

        prioridad = "🟡 MEDIA";

    }

    mensaje += "<h3>📊 Diagnóstico automático</h3>";

    mensaje += `<p>Se analizaron <strong>${total}</strong> estudiantes.</p>`;

    mensaje += `<p>Se detectaron <strong>${alto}</strong> estudiantes con riesgo alto.</p>`;

    mensaje += `<p>El indicador más crítico es <strong>${indicadorCritico}</strong>.</p>`;

    mensaje += "<hr>";

    mensaje += "<h3>🤖 Recomendaciones</h3>";

    if(indicadorCritico=="Asistencia"){

        mensaje += "<p>✔ Implementar seguimiento semanal de asistencia.</p>";

        mensaje += "<p>✔ Reuniones con padres de familia.</p>";

        mensaje += "<p>✔ Incentivar la asistencia mediante actividades.</p>";

    }

    if(indicadorCritico=="Promedio"){

        mensaje += "<p>✔ Tutorías académicas.</p>";

        mensaje += "<p>✔ Reforzamiento después de clases.</p>";

        mensaje += "<p>✔ Evaluaciones diagnósticas.</p>";

    }

    if(indicadorCritico=="Tareas"){

        mensaje += "<p>✔ Seguimiento individual.</p>";

        mensaje += "<p>✔ Cronograma de entregas.</p>";

        mensaje += "<p>✔ Comunicación permanente con las familias.</p>";

    }

    mensaje += "<hr>";

    mensaje += `<h3>Nivel de prioridad: ${prioridad}</h3>`;

    resultadoIA.innerHTML = mensaje;

}

botonTodos.addEventListener("click",function(){

    filtrarPorRiesgo("TODOS");

});

botonAlto.addEventListener("click",function(){

    filtrarPorRiesgo("ALTO");

});

botonMedio.addEventListener("click",function(){

    filtrarPorRiesgo("MEDIO");

});

botonBajo.addEventListener("click",function(){

    filtrarPorRiesgo("BAJO");

});

const botonPDF = document.getElementById("exportarPDF");

botonPDF.addEventListener("click", generarPDF);

function generarPDF(){

    const { jsPDF } = window.jspdf;

    const doc = new jsPDF();

    //=========================
    // DATOS
    //=========================

    const colegio = document.getElementById("nombreColegio").value || "No especificado";

    const docente = document.getElementById("nombreDocente").value || "No especificado";

    const fecha = new Date();

    const fechaTexto = fecha.toLocaleDateString();

    const horaTexto = fecha.toLocaleTimeString();

    let resumen = document.getElementById("resumenAnalisis").innerText;

    let recomendaciones = document.getElementById("resultadoIA").innerText;

    //=========================
    // LIMPIAR EMOJIS
    //=========================

    const limpiar = (texto)=>{

        return texto
        .replace(/[^\x00-\x7F]/g,"")
        .replace(/\s+/g," ")
        .trim();

    }

    resumen = limpiar(resumen);

    recomendaciones = limpiar(recomendaciones);

    //=========================
    // ENCABEZADO
    //=========================

    doc.setFillColor(25,118,210);

    doc.rect(0,0,210,28,"F");

    doc.setTextColor(255);

    doc.setFont("helvetica","bold");

    doc.setFontSize(18);

    doc.text("SISTEMA INTELIGENTE DE ANALITICA EDUCATIVA",105,12,{align:"center"});

    doc.setFontSize(11);

    doc.text("Informe generado automaticamente",105,20,{align:"center"});

    //=========================
    // COLOR NORMAL
    //=========================

    doc.setTextColor(0);

    let y = 40;

    doc.setFontSize(12);

    doc.setFont("helvetica","bold");

    doc.text("DATOS GENERALES",20,y);

    y+=10;

    doc.setFont("helvetica","normal");

    doc.text(`Colegio: ${colegio}`,20,y);

    y+=8;

    doc.text(`Docente: ${docente}`,20,y);

    y+=8;

    doc.text(`Fecha: ${fechaTexto}`,20,y);

    y+=8;

    doc.text(`Hora: ${horaTexto}`,20,y);

    y+=15;

    //=========================
    // RESUMEN
    //=========================

    doc.setFont("helvetica","bold");

    doc.text("RESUMEN DEL ANALISIS",20,y);

    y+=10;

    doc.setFont("helvetica","normal");

    const resumenPDF = doc.splitTextToSize(resumen,170);

    doc.text(resumenPDF,20,y);

    y += resumenPDF.length*7+10;

    //=========================
    // RECOMENDACIONES
    //=========================

    doc.setFont("helvetica","bold");

    doc.text("RECOMENDACIONES INTELIGENTES",20,y);

    y+=10;

    doc.setFont("helvetica","normal");

    const recomendacionesPDF = doc.splitTextToSize(recomendaciones,170);

    doc.text(recomendacionesPDF,20,y);

    //=========================
    // NOTA FINAL
    //=========================

    y += recomendacionesPDF.length * 6 + 10;

    doc.setDrawColor(180);

    doc.line(20, y, 190, y);

    y += 8;

    doc.setFont("helvetica","italic");

    doc.setFontSize(10);

    doc.setTextColor(80);

    const nota = doc.splitTextToSize(

        "Nota: Para una mejor visualizacion de los graficos, panel interactivo y analisis detallado de cada estudiante, consulte el Sistema Inteligente de Analitica Educativa desde la aplicacion web.",

        170

    );

    doc.text(nota,20,y);

    y += nota.length * 5 + 10;

    //=========================
    // PIE
    //=========================

    doc.setFont("helvetica","normal");

    doc.setFontSize(9);

    doc.setTextColor(120);

    doc.text(

        "Documento generado automaticamente por el Sistema Inteligente de Analitica Educativa.",

        105,

        285,

        {align:"center"}

    );
    doc.save("Reporte_Analitica_Educativa.pdf");

}