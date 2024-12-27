const path = require('path');
const XLSX = require('xlsx-js-style');

const rf = config.tableType == "ods" ? '.' : '!'
const weeksSheetName = "Avancé journalière"

const white = "FFFFFF"; // #FFFFFF

const emptyStyle = {fill : {fgColor : { rgb: white }}} // White color (RGB code for white)
const emptyCell = { v: '',
    // s: emptyStyle
}; // Empty cell with white background

const dateColor = "d8d8d8" // #d8d8d8

const gradeColors = {
    6 : 'C6233D', // #C6233D
    5 : '088255', // #088255
    4 : '1466A8', // #1466A8
    3 : '844499'  // #844499
}

const weeks = require(path.join(__basedir,config.input,'weeks.js'));

const borderStyle = { style: "thin", color: { rgb: "000000" }};

const weekHeaderStyle = {
    alignment: { horizontal: 'center', vertical: 'center' }, // Centered text
    font: { bold: true } // Bold text
};

function addClassHeader(c,cols,row){
    const gradeHeaderStyle = {
        alignment: { horizontal: 'center', vertical: 'center' }, // Centered text
        font: { bold: true , color : { rgb: gradeColors[c.name[Number(0)]] } } // Bold text
    };
    row.push({ v: c.name, s: gradeHeaderStyle },emptyCell,emptyCell)
    cols.push({ wch: 11 },{ wch: 30 },{ wch: 2 })
}

function generateWeeksSheet(classes) {
    const content = []; // Tableau pour stocker les données
    const cols = [{ wch: 2 },{ wch: 8 }, { wch: 10.5 },{ wch: 2 }]
    // const merges = []
    const firstRow = [emptyCell,{ v: `Semaines`, s: weekHeaderStyle },emptyCell,emptyCell]

    // let maxHoures = 0

    // const days = ['monday','tuesday','wednesday','thursday','friday','saturday']
    // days.forEach(day => {
    //     classes.forEach(c => {
    //         if(c[day]){
    //             maxHoures = Math.max(c[day].class.length,maxHoures)
    //         }
    //     });
    // });

    // console.log(maxHoures)

    classes.forEach(c => {
        addClassHeader(c,cols,firstRow)
    });

    content.push(firstRow)

    // Génération des données hebdomadaires
    for (let week = weeks.first; week <= weeks.last; week++) {
        content.push([emptyCell,{ v: `${week}`, s: weekHeaderStyle }, emptyCell]);
        if(!weeks.hollidays.includes(week)){
            for (let dayOffset = 0; dayOffset < weeks.length; dayOffset++) { // Du lundi au vendredi
                const day = getDay(week,dayOffset,weeks.startDate)
                addDay(content,classes,day)
            }
        }
    }

    // Création de la feuille de calcul
    const sheet = XLSX.utils.aoa_to_sheet(content);

    sheet["!merges"] = [];
    
    let rowIndex = 0;
    sheet["!merges"].push({
        s: { r: rowIndex, c: 1 }, // Début de la fusion (ligne rowIndex, colonne 0)
        e: { r: rowIndex, c: 2 }  // Fin de la fusion (ligne rowIndex, colonne 1)
    });

    colIndex = 4;
    classes.forEach(c => {
        sheet["!merges"].push({
            s: { r: rowIndex, c: colIndex }, // Début de la fusion (ligne rowIndex, colonne 0)
            e: { r: rowIndex, c: colIndex+1 }  // Fin de la fusion (ligne rowIndex, colonne 1)
        });
        colIndex+=3;
    });

    rowIndex++;

    for (let week = weeks.first; week <= weeks.last; week++) {
        sheet["!merges"].push({
            s: { r: rowIndex, c: 1 }, // Début de la fusion (ligne rowIndex, colonne 0)
            e: { r: rowIndex, c: 2 }  // Fin de la fusion (ligne rowIndex, colonne 1)
        });
        rowIndex += weeks.hollidays.includes(week)? 1 : 6; // Une ligne pour l'en-tête + 5 jours
    }

    sheet["!cols"] = cols

    return sheet;
}

function getDay(week,dayOffset,startDate) {
    const currentDate = new Date(startDate);
    currentDate.setDate(startDate.getDate() + (week - 1) * 7 + dayOffset);
    day = {
        week : week,
        id : dayOffset,
        date : currentDate,
        name : currentDate.toLocaleDateString("en", {weekday: "long" }).toLocaleLowerCase(),
        fr : currentDate.toLocaleDateString("fr-FR", { weekday: "long" }),
        formattedDate : currentDate.toLocaleDateString("fr-FR")
    }
    return day;
}

function addDay(content,classes,day) {
    // console.log(formattedDate)

    const dayStyle = {
        alignment: { horizontal: 'left', vertical: 'center' }, // Centered text
        fill: { // Background color (light blue in this case)
            fgColor: { rgb: fun.blendColors(white, dateColor,day.id % 2 == 0 ? 0.5 : 1) }
        },
        border: {
            top: (day.id == 0)?  borderStyle : null,
            bottom: (day.id == weeks.length - 1)? borderStyle : null,
            left: borderStyle
        }
    };

    const dateStyle = {
        alignment: { horizontal: 'center', vertical: 'center' }, // Centered text
        fill: { // Background color (light blue in this case)
            fgColor: { rgb: fun.blendColors(white, dateColor,day.id % 2 == 0 ? 0.15 : 0.45) }
        },
        border: {
            top: (day.id == 0)?  borderStyle : null,
            bottom: (day.id == weeks.length - 1)? borderStyle : null,
            right: borderStyle
        }
    };

    day.row  = [emptyCell,{ v: day.fr, s: dayStyle }, { v: day.formattedDate, s: dateStyle},emptyCell]
    classes.forEach(c => {
        addClass(c,day)
    });
    content.push(day.row); // Ajoute jour et date dans la ligne
}

function addClass(c,day) {
    const defaultGradeColor = gradeColors[c.name[Number(0)]]
    const gradeColor = fun.blendColors(white, defaultGradeColor,day.id % 2 == 0 ? 0.75 : 1);

    const emptyStyle = { rgb: fun.blendColors(white, defaultGradeColor, 0.15) }
    
    const gradeStyle = {
        fill: { // Background color (light blue in this case)
            fgColor: c[day.name]? { rgb: fun.blendColors(white, gradeColor, 0.5) } : emptyStyle // Light blue color (RGB code for light blue)
        },
        border: {
            top: (day.id == 0)?  borderStyle : null,
            bottom: (day.id == weeks.length - 1)? borderStyle : null,
            right: borderStyle
        }
    };

    const gradeTimeStyle = {
        alignment: {horizontal: 'right', vertical: 'center' },
        fill: { // Background color (light blue in this case)
            fgColor: c[day.name]? { rgb: gradeColor } : emptyStyle // Light blue color (RGB code for light blue)
        },
        border: {
            top: (day.id == 0)?  borderStyle : null,
            bottom: (day.id == weeks.length - 1)? borderStyle : null,
            left: borderStyle
        }
    };
    // classStyle.fill.fgColor.rgb = currentColor;
    if(c[day.name]){
        const times = []
        const coments = []
        for (let h = 0; h < c[day.name].class.length; h++) {
            let t = c[day.name].class[h]

            if(t.week == null
                || (t.week === 0 && day.week % 2 == 0)
                || (t.week === 1 && day.week % 2 == 1)){
                    let time = String(t.time);
                    if (time.includes('.')) {
                        const [hours, minutes] = time.split('.'); // Split into hours and minutes
                        time = `${hours.padStart(2, '0')}:${minutes.padEnd(2, '0')}`; // Format as HH:MM
                    }else{
                        time = `${time.padStart(2, '0')}:00`;
                    }
                    times.push(time)
                    coments.push(`(${t.coment})` ? t.coment : emptyCell)
            }
        }
        day.row.push({ v: times.join(' / '), s: gradeTimeStyle },{ v: coments.join(' / '), s: gradeStyle }) 
    }else{
        day.row.push({ v: '', s: gradeTimeStyle },{ v: '', s: gradeStyle })
    }
    day.row.push(emptyCell)
}

function writeSheet() {
    const classes = fun.loadClassesSchedules();
    const workbook = XLSX.utils.book_new();
    const weeksSheet = generateWeeksSheet(classes);
    XLSX.utils.book_append_sheet(workbook, weeksSheet, weeksSheetName);
    XLSX.writeFile(workbook, config.output + "." + config.tableType , { bookType: config.tableType });
    console.log(`Sheet written at : ${config.output}.${config.tableType}`);
}

module.exports = () => {
    writeSheet();
};