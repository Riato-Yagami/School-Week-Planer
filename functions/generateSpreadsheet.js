const path = require('path');
const fs = require('fs');
const { strFromU8, strToU8, unzipSync, zipSync } = require('fflate');
const XLSX = require('xlsx-js-style');

const rf = config.tableType == "ods" ? '.' : '!'
const weeklySheetName = "Emploi du temps semaine"
const weeksSheetName = "Avancé journalière"

const colors = require(path.join(__basedir, 'config/colors'));
const schoolHours = require(path.join(__basedir, 'config/school-hours'));
const white = colorHex(colors.white);

const emptyStyle = {fill : {fgColor : { rgb: white }}} // White color (RGB code for white)
const emptyCell = { v: '',
    // s: emptyStyle
}; // Empty cell with white background

const dateColor = colorHex(colors.date);
const todayColor = colorHex(colors.today);
const gradeColors = Object.fromEntries(Object.entries(colors.grades).map(([grade, color]) => [grade, colorHex(color)]));

const weeks = require(path.join(__basedir,config.input,'weeks.js'));
let previousComments = new Map();
let currentPdfPageHeight = 595;

const borderStyle = { style: "thin", color: { rgb: colorHex(colors.text) }};

const weekHeaderStyle = {
    alignment: { horizontal: 'center', vertical: 'center' }, // Centered text
    font: { bold: true } // Bold text
};

const summaryHeaderStyle = {
    alignment: { horizontal: 'center', vertical: 'center' },
    font: { bold: true },
    fill: { fgColor: { rgb: colorHex(colors.summaryHeader) } },
    border: {
        top: borderStyle,
        bottom: borderStyle,
        left: borderStyle,
        right: borderStyle
    }
};

const summaryCellStyle = {
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
        top: borderStyle,
        bottom: borderStyle,
        left: borderStyle,
        right: borderStyle
    }
};

const timetableHeaderStyle = {
    alignment: { horizontal: 'center', vertical: 'center' },
    font: { bold: true },
    fill: { fgColor: { rgb: dateColor } },
    border: {
        top: borderStyle,
        bottom: borderStyle,
        left: borderStyle,
        right: borderStyle
    }
};

const timetableEmptyStyle = {
    alignment: { horizontal: 'center', vertical: 'center' },
    font: { color: { rgb: colorHex(colors.commentText) } },
    fill: { fgColor: { rgb: colorHex(colors.timetableEmpty) } },
    border: {
        top: borderStyle,
        bottom: borderStyle,
        left: borderStyle,
        right: borderStyle
    }
};

const offLabelStyle = {
    alignment: { horizontal: 'center', vertical: 'center' },
    font: { bold: true, italic: true, color: { rgb: colorHex(colors.commentText) } }
};

function addClassHeader(c,cols,row){
    const gradeHeaderStyle = {
        alignment: { horizontal: 'center', vertical: 'center' }, // Centered text
        font: { bold: true , color : { rgb: gradeColors[c.name[Number(0)]] } }, // Bold text
        border: {
            top: borderStyle,
            bottom: borderStyle,
            left: borderStyle,
            right: borderStyle
        }
    };
    row.push({ v: formatClassName(c.name), s: gradeHeaderStyle },{ v: '', s: gradeHeaderStyle },emptyCell)
    cols.push({ wch: 11 },{ wch: 30 },{ wch: 2 })
}

function generateWeeksSheet(classes) {
    const content = []; // Tableau pour stocker les données
    const cols = [{ wch: 2 },{ wch: 8 }, { wch: 10.5 },{ wch: 2 }]
    // const merges = []
    const firstRow = [emptyCell,{ v: `Semaines`, s: weekHeaderStyle },{ v: '', s: weekHeaderStyle },emptyCell]

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
    cols.push({ wch: 11 },{ wch: 30 },{ wch: 2 })

    content.push(firstRow)

    initializeOffDays();

    // console.log(weeks.offDays)
    // Génération des données hebdomadaires
    for (let week = weeks.first; week <= weeks.last; week++) {
        if(weeks.holidays.includes(week)){
            addHolidayWeek(content, classes, week);
        }else{
            content.push([emptyCell,{ v: `${week}`, s: weekHeaderStyle }, { v: '', s: weekHeaderStyle }, emptyCell]);
        }
        if(!weeks.holidays.includes(week)){
            for (let dayOffset = 0; dayOffset < weeks.length; dayOffset++) { // Du lundi au vendredi
                const day = getDay(week,dayOffset,weeks.startDate);
                // console.log(day.formattedDate)
                // if(weeks.formattedOffDays.includes(day.formattedDate)){
                //     console.log(`Off day : ${day.formattedDate} (${day.fr})`)
                //     addDay(content,[],day)
                // }else{
                    addDay(content,classes,day)
                // }
            }
        }
    }

    // Création de la feuille de calcul
    const offLabelGroups = normalizeOffLabelGroups(content);
    const summaryStartRow = content.length;
    addHoursSummary(content, classes, summaryStartRow)

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
        if (weeks.holidays.includes(week)) {
            rowIndex += 1;
        } else {
            rowIndex++;
            for (let dayOffset = 0; dayOffset < weeks.length; dayOffset++) {
                rowIndex++;
            }
        }
    }

    addOffLabelGroupMerges(sheet, classes, offLabelGroups)
    addHoursSummaryMerges(sheet, classes, summaryStartRow)

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

function initializeOffDays() {
    const offDayMap = new Map();
    const offDayLabels = Object.assign({}, weeks.offDayLabels || {});

    (weeks.offDays || []).forEach(date => {
        offDayMap.set(formatDateKey(date), date);
    });

    weeks.offDays = Array.from(offDayMap.values()).sort((a, b) => a - b);
    weeks.offDayLabels = offDayLabels;
    weeks.formattedOffDays = Array.from(offDayMap.keys());
    weeks.holidays = Array.from(new Set([...(weeks.holidays || []), ...getFullOffWeeks()])).sort((a, b) => a - b);
}

function formatDateKey(date) {
    return date.toLocaleDateString("fr-FR");
}

function getFullOffWeeks() {
    const holidays = [];
    const offDayKeys = new Set(weeks.formattedOffDays);

    for (let week = weeks.first; week <= weeks.last; week++) {
        let fullOffWeek = true;
        for (let dayOffset = 0; dayOffset < weeks.length; dayOffset++) {
            const day = getDay(week, dayOffset, weeks.startDate);
            fullOffWeek = fullOffWeek && offDayKeys.has(day.formattedDate);
        }

        if (fullOffWeek) {
            holidays.push(week);
        }
    }

    return holidays;
}

function addDay(content,classes,day) {
    // console.log(formattedDate)

    const dayStyle = {
        alignment: { horizontal: 'left', vertical: 'center' }, // Centered text
        fill: { // Background color (light blue in this case)
            fgColor: { rgb: fun.blendColors(white, dateColor, day.id % 2 == 0 ? 0.5 : 1) }
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

    day.row  = [emptyCell,{ v: day.fr, s: dayStyle }, { v: day.date, t: 'd', z: 'dd/mm/yyyy', s: dateStyle},emptyCell]
    if (isOffDay(day)) {
        addOffLabelCells(day.row, classes, getOffDayLabel(day));
    } else {
        classes.forEach(c => {
            addClass(c,day)
        });
    }
    content.push(day.row); // Ajoute jour et date dans la ligne
}

function addHolidayWeek(content, classes, week) {
    const row = [emptyCell,{ v: `${week}`, s: weekHeaderStyle }, { v: '', s: weekHeaderStyle }, emptyCell];
    addOffLabelCells(row, classes, getHolidayWeekLabel(week));
    content.push(row);
}

function addOffLabelCells(row, classes, label) {
    for (let col = row.length; col <= getLastClassCommentCol(classes); col++) {
        row.push({ v: '', s: offLabelStyle });
    }

    row[4] = { v: label, s: offLabelStyle };
}

function normalizeOffLabelGroups(content) {
    const groups = [];
    let currentGroup = null;

    content.forEach((row, rowIndex) => {
        const labelCell = row[4];
        const label = labelCell && labelCell.s === offLabelStyle ? String(labelCell.v || '') : '';

        if (!label) {
            currentGroup = null;
            return;
        }

        if (!currentGroup || currentGroup.label !== label || currentGroup.endRow !== rowIndex - 1) {
            currentGroup = { label, startRow: rowIndex, endRow: rowIndex };
            groups.push(currentGroup);
            return;
        }

        currentGroup.endRow = rowIndex;
        labelCell.v = '';
    });

    return groups;
}

function addOffLabelGroupMerges(sheet, classes, groups) {
    if (!classes.length) {
        return;
    }

    groups.forEach(group => {
        sheet["!merges"].push({
            s: { r: group.startRow, c: 4 },
            e: { r: group.endRow, c: getLastClassCommentCol(classes) }
        });
    });
}

function getLastClassCommentCol(classes) {
    return 5 + Math.max(0, classes.length - 1) * 3;
}

function isOffDay(day) {
    return weeks.formattedOffDays.includes(day.formattedDate);
}

function getOffDayLabel(day) {
    return getConfiguredOffDayLabel(day.formattedDate) || 'Férié';
}

function getHolidayWeekLabel(week) {
    for (let dayOffset = 0; dayOffset < weeks.length; dayOffset++) {
        const day = getDay(week, dayOffset, weeks.startDate);
        const label = getConfiguredOffDayLabel(day.formattedDate);
        if (label) {
            return label;
        }
    }

    return 'Vacances';
}

function getConfiguredOffDayLabel(formattedDate) {
    return weeks.offDayLabels && weeks.offDayLabels[formattedDate];
}

function goodWeek(time,day){
    return (time.week === 0 && day.week % 2 == 0)
    || (time.week === 1 && day.week % 2 == 1)
}

function shouldShowTime(time, day) {
    if (time.date != null) {
        return time.date === day.formattedDate;
    }

    if (time.week != null) {
        return goodWeek(time, day);
    }

    return true;
}

function getShownTimes(c, day) {
    if (!c[day.name]) {
        return [];
    }

    return c[day.name].class.filter(time => shouldShowTime(time, day));
}

function getCommentKey(className, day) {
    return `${className}|${day.formattedDate}`;
}

function getCommentFallbackKey(className, day) {
    return `${className}|week:${day.week}|day:${day.id}`;
}

function getPreviousComment(className, day) {
    return previousComments.get(getCommentKey(className, day))
        || previousComments.get(getCommentFallbackKey(className, day))
        || '';
}

function splitCommentSegments(comment) {
    return String(comment || '').split(/\s+\/\s+/);
}

function hasManualComment(comment, lessonIndex) {
    const segments = splitCommentSegments(comment);
    const segment = segments[lessonIndex] != null ? segments[lessonIndex] : '';
    const manualText = segment
        .replace(/\[[^\]]+\]/g, '')
        .replace(/\//g, '')
        .trim();

    return manualText.length > 0;
}

function getManualComment(comment, lessonIndex) {
    const segments = splitCommentSegments(comment);
    const segment = segments[lessonIndex] != null ? segments[lessonIndex] : '';

    return segment
        .replace(/\[[^\]]+\]/g, '')
        .replace(/\//g, '')
        .trim();
}

function formatCommentWithManual(c, day, time, lessonIndex) {
    const autoComment = formatComment(time.comment);
    const previousComment = getPreviousComment(c.name, day);
    const manualComment = getManualComment(previousComment, lessonIndex);

    return `${autoComment}${manualComment ? ` ${manualComment}` : ''}`.trim();
}

function getLessonDuration(time) {
    return time.duration != null ? time.duration : 1;
}

// Excel stores durations as fractions of a day; elapsed hours may exceed 24.
const hoursNumberFormat = '[h]"h"mm"min"';

function getSummaryStats(content, classes) {
    const durationsByRow = new Map();
    content.forEach((row, index) => {
        if (!row[2] || row[2].t !== 'd') return;
        const date = row[2].v;
        const day = {
            date, formattedDate: formatDateKey(date),
            name: date.toLocaleDateString('en', { weekday: 'long' }).toLowerCase(),
            week: Math.floor((Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
                - Date.UTC(weeks.startDate.getFullYear(), weeks.startDate.getMonth(), weeks.startDate.getDate())) / 86400000 / 7) + 1
        };
        if (isOffDay(day)) return;
        const times = classes.flatMap(c => getShownTimes(c, day));
        if (times.length) durationsByRow.set(index + 1, {
            date, hours: times.reduce((sum, time) => sum + getLessonDuration(time), 0)
        });
    });
    const today = new Date();
    today.setHours(23, 59, 59, 999);
    const entries = Array.from(durationsByRow.entries());
    return {
        total: entries.reduce((sum, [, value]) => sum + value.hours, 0) / 24,
        done: entries.reduce((sum, [, value]) => sum + (value.date <= today ? value.hours : 0), 0) / 24,
        daysTotal: entries.length,
        daysDone: entries.filter(([, value]) => value.date <= today).length,
        doneFormula: sumFormula(entries.map(([row, value]) => `IF($C${row}<=TODAY(),${value.hours}/24,0)`)),
        daysDoneFormula: sumFormula(entries.map(([row]) => `IF($C${row}<=TODAY(),1,0)`))
    };
}

function addHoursSummary(content, classes, summaryStartRow) {
    const labels = ['', 'Heures', 'Faites', 'Total', 'Restant', 'Complétion', '', 'Jours', 'Faits', 'Total', 'Restant'];
    const rows = labels.map((label, index) => index === 0 ? [emptyCell] :
        [emptyCell, { v: label, s: summaryHeaderStyle }, { v: '', s: summaryHeaderStyle }, emptyCell]);
    rows[6] = [emptyCell, emptyCell, emptyCell, emptyCell];
    [...classes, null].forEach((c, index) => {
        // The overall day count counts each teaching date once, even with several classes.
        const stats = getSummaryStats(content, c ? [c] : classes);
        const col = 4 + index * 3;
        const style = c ? summaryCellStyle : summaryHeaderStyle;
        const ref = offset => getCellRef(summaryStartRow + offset, col);
        const classHeader = { v: c ? formatClassName(c.name) : 'Total', s: summaryHeaderStyle };
        const values = [
            formulaCell(stats.doneFormula, stats.done, style, hoursNumberFormat),
            formulaCell(String(stats.total), stats.total, style, hoursNumberFormat),
            formulaCell(`${ref(3)}-${ref(2)}`, stats.total - stats.done, style, hoursNumberFormat),
            formulaCell(`IF(${ref(3)}=0,0,${ref(2)}/${ref(3)})`, stats.total ? stats.done / stats.total : 0, style, '0.0%'),
            formulaCell(stats.daysDoneFormula, stats.daysDone, style),
            formulaCell(String(stats.daysTotal), stats.daysTotal, style),
            formulaCell(`${ref(10)}-${ref(9)}`, stats.daysTotal - stats.daysDone, style)
        ];
        setRowCell(rows[1], col, classHeader);
        setRowCell(rows[1], col + 1, { v: '', s: classHeader.s });
        values.slice(0, 4).forEach((cell, offset) => {
            setRowCell(rows[offset + 2], col, cell);
            setRowCell(rows[offset + 2], col + 1, { v: '', s: cell.s });
        });
        values.slice(4).forEach((cell, offset) => {
            setRowCell(rows[offset + 8], col, cell);
            setRowCell(rows[offset + 8], col + 1, { v: '', s: cell.s });
        });
    });
    content.push(...rows);
}

function formulaCell(formula, value, style, format = '0') {
    return { t: 'n', f: formula, v: value, z: format, s: style };
}

function sumFormula(refs) {
    return refs.length ? `SUM(${refs.join(',')})` : '0';
}

function getCellRef(rowIndex, colIndex) {
    return XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
}

function addHoursSummaryMerges(sheet, classes, summaryStartRow) {
    [1, 2, 3, 4, 5, 7, 8, 9, 10].forEach(rowOffset => {
        sheet["!merges"].push({
            s: { r: summaryStartRow + rowOffset, c: 1 },
            e: { r: summaryStartRow + rowOffset, c: 2 }
        });

        for (let index = 0; index <= classes.length; index++) {
            const colIndex = 4 + index * 3;
            sheet["!merges"].push({
                s: { r: summaryStartRow + rowOffset, c: colIndex },
                e: { r: summaryStartRow + rowOffset, c: colIndex + 1 }
            });
        }
    });
}

function setRowCell(row, index, cell) {
    while (row.length < index) {
        row.push(emptyCell);
    }

    row[index] = cell;
}

function formatClassName(className) {
    return String(className).trim().replace(/^(\d)\s+(\d+)$/, '$1e$2');
}

function addClass(c,day) {
    let workableDay = (c[day.name] && !weeks.formattedOffDays.includes(day.formattedDate))
    
    if(workableDay){
        workableDay = false
        c[day.name].class.forEach(t => {
            if(shouldShowTime(t, day)){
                workableDay = true
            }
        })
    }

    const defaultGradeColor = gradeColors[c.name[Number(0)]]
    const gradeColor = fun.blendColors(white, defaultGradeColor,day.id % 2 == 0 ? 0.75 : 1);

    const emptyStyle = { rgb: fun.blendColors(white, defaultGradeColor, 0.05) }
    
    const gradeStyle = {
        font: { color: { rgb: colorHex(colors.text) } },
        fill: {
            fgColor: workableDay
                ? { rgb: fun.blendColors(white, gradeColor, 0.35) }
                : emptyStyle
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
            fgColor: workableDay? { rgb: gradeColor } : emptyStyle // Light blue color (RGB code for light blue)
        },
        border: {
            top: (day.id == 0)?  borderStyle : null,
            bottom: (day.id == weeks.length - 1)? borderStyle : null,
            left: borderStyle
        }
    };
    // classStyle.fill.fgColor.rgb = currentColor;
    if(workableDay){
        const times = []
        const comments = []
        const shownTimes = getShownTimes(c, day)
        for (let h = 0; h < shownTimes.length; h++) {
            let t = shownTimes[h]

            let time = String(t.time);
            if (time.includes('.')) {
                const [hours, minutes] = time.split('.'); // Split hours and minutes
                time = `${hours.padStart(2, '0')}:${minutes.padEnd(2, '0')}`;
            }else{
                time = `${time.padStart(2, '0')}:00`;
            }
            times.push(time)
            comments.push(formatCommentWithManual(c, day, t, h))
        }
        const displayedComments = comments;
        day.row.push({ v: times.join(' / '), s: gradeTimeStyle },{ v: displayedComments.join(' / '), s: gradeStyle }) 
    }else{
        day.row.push({ v: '', s: gradeTimeStyle },{ v: '', s: gradeStyle })
    }
    day.row.push(emptyCell)
}

function formatComment(comment) {
    if (!comment) {
        return '';
    }

    const shortComment = String(comment)
        .replace(/\bVIE DE CLASSE\b/gi, 'vdc')
        .replace(/\b\d\s*\d?_?P(\d)\b(?:\s+non euro)?/gi, 'g$1')
        .replace(/\s*-\s*/g, ' - ')
        .trim();

    return shortComment ? `[${shortComment}]` : '';
}

function loadPreviousComments(classes, outputPath) {
    const comments = new Map();
    const days = ['lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi'];
    if (!fs.existsSync(outputPath)) {
        return comments;
    }

    const workbook = XLSX.readFile(outputPath);
    const sheet = workbook.Sheets[weeksSheetName] || workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet || !sheet["!ref"]) {
        return comments;
    }

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    classes.forEach((c, index) => {
        const commentCol = 5 + index * 3;
        let currentWeek = null;

        for (let row = range.s.r; row <= range.e.r; row++) {
            const dayOrWeekCell = sheet[XLSX.utils.encode_cell({ r: row, c: 1 })];
            const dateCell = sheet[XLSX.utils.encode_cell({ r: row, c: 2 })];
            const commentCell = sheet[XLSX.utils.encode_cell({ r: row, c: commentCol })];
            const dayOrWeek = getCellText(dayOrWeekCell).trim();
            const date = getCellText(dateCell).trim();

            if (/^\d+$/.test(dayOrWeek) && !date) {
                currentWeek = Number(dayOrWeek);
            }

            if (!dateCell || !commentCell || !commentCell.v) {
                continue;
            }

            comments.set(`${c.name}|${date}`, String(commentCell.v));

            const dayId = days.indexOf(dayOrWeek.toLocaleLowerCase('fr-FR'));
            if (currentWeek != null && dayId !== -1) {
                comments.set(`${c.name}|week:${currentWeek}|day:${dayId}`, String(commentCell.v));
            }
        }
    });

    return comments;
}

function getCellText(cell) {
    if (!cell) {
        return '';
    }

    return cell.w != null ? String(cell.w) : String(cell.v || '');
}

function colorizeCommentCells(outputPath, classes) {
    if (config.tableType !== 'xlsx' || !fs.existsSync(outputPath)) {
        return;
    }

    const zip = unzipSync(fs.readFileSync(outputPath));
    Object.keys(zip)
        .filter(filePath => filePath.startsWith('xl/worksheets/') && filePath.endsWith('.xml'))
        .forEach(filePath => {
            const xml = strFromU8(zip[filePath]);
            const updatedXml = xml.replace(/<c([^>]*) t="str"([^>]*)><v([^>]*)>([\s\S]*?)<\/v><\/c>/g, (match, beforeType, afterType, valueAttributes, value) => {
                const text = xmlUnescape(value);
                if (!needsRichText(text)) {
                    return match;
                }

                return `<c${beforeType} t="inlineStr"${afterType}><is>${toRichTextXml(text)}</is></c>`;
            });

            zip[filePath] = strToU8(updatedXml);
        });

    const sharedStringsPath = 'xl/sharedStrings.xml';
    if (zip[sharedStringsPath]) {
        const xml = strFromU8(zip[sharedStringsPath]);
        const updatedXml = xml.replace(/<si><t(?: xml:space="preserve")?>([\s\S]*?)<\/t><\/si>/g, (match, value) => {
            const text = xmlUnescape(value);
            if (!needsRichText(text)) {
                return match;
            }

            return `<si>${toRichTextXml(text)}</si>`;
        });

        zip[sharedStringsPath] = strToU8(updatedXml);
    }

    addTodayConditionalFormatting(zip, classes);
    fs.writeFileSync(outputPath, Buffer.from(zipSync(zip)));
}

function addTodayConditionalFormatting(zip, classes) {
    const stylesPath = 'xl/styles.xml';
    const sheetPath = 'xl/worksheets/sheet1.xml';

    if (!zip[stylesPath] || !zip[sheetPath]) {
        return;
    }

    const borderId = addTodayDxfStyle(zip, stylesPath, `<border>${['left', 'right', 'top', 'bottom'].map(side => `<${side} style="medium"><color rgb="FF${todayColor}"/></${side}>`).join('')}</border>`);
    const highlightId = addTodayDxfStyle(zip, stylesPath,
        `<fill><patternFill patternType="solid"><fgColor rgb="FF${todayColor}"/><bgColor rgb="FF${todayColor}"/></patternFill></fill>` +
        `<border>${['left', 'right', 'top', 'bottom'].map(side => `<${side} style="medium"><color rgb="FF${todayColor}"/></${side}>`).join('')}</border>`);
    const xml = strFromU8(zip[sheetPath]);
    const rules = [];
    const endCol = XLSX.utils.encode_col(getLastClassCommentCol(classes));
    let row = 2;
    const addRule = (range, formula, style) => rules.push(
        `<conditionalFormatting sqref="${range}"><cfRule type="expression" dxfId="${style}" priority="${rules.length + 1}"><formula>${formula}</formula></cfRule></conditionalFormatting>`);
    for (let week = weeks.first; week <= weeks.last; week++) {
        const monday = getDay(week, 0, weeks.startDate).date;
        const dateFormula = `DATE(${monday.getFullYear()},${monday.getMonth() + 1},${monday.getDate()})`;
        // The header of the coming week marks the weekend between two weeks.
        addRule(`B${row}:${endCol}${row}`, `AND(TODAY()&gt;=${dateFormula}-2,TODAY()&lt;${dateFormula},WEEKDAY(TODAY(),2)&gt;5)`, borderId);
        row += weeks.holidays.includes(week) ? 1 : weeks.length + 1;
    }
    addRule(`B1:C${row - 1}`, 'AND(ISNUMBER($C1),INT($C1)=TODAY())', highlightId);
    for (let index = 0; index < classes.length; index++) {
        const col = 4 + index * 3;
        const separatorCol = XLSX.utils.encode_col(col - 1);
        const timeCol = XLSX.utils.encode_col(col);
        addRule(`${separatorCol}1:${timeCol}${row - 1}`,
            'AND(ISNUMBER($C1),INT($C1)=TODAY())', highlightId);
        addRule(`${XLSX.utils.encode_col(col)}1:${XLSX.utils.encode_col(col + 1)}${row - 1}`,
            'AND(ISNUMBER($C1),INT($C1)=TODAY())', borderId);
    }
    // Also mark the weekend following the final week of the planning.
    const end = getDay(weeks.last + 1, 0, weeks.startDate).date;
    const endFormula = `DATE(${end.getFullYear()},${end.getMonth() + 1},${end.getDate()})`;
    addRule(`B${row - 1}:${endCol}${row - 1}`, `AND(TODAY()&gt;=${endFormula}-2,TODAY()&lt;${endFormula},WEEKDAY(TODAY(),2)&gt;5)`, borderId);
    // Conditional formatting precedes page/print settings in the worksheet schema.
    const insertAt = /<(?:dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|drawing|extLst)\b|<\/worksheet>/;
    zip[sheetPath] = strToU8(xml.replace(insertAt, match => rules.join('') + match));
    const workbookPath = 'xl/workbook.xml';
    if (zip[workbookPath]) {
        const workbook = strFromU8(zip[workbookPath]).replace(/<calcPr[^>]*\/>/g, '');
        zip[workbookPath] = strToU8(workbook.replace('</workbook>', '<calcPr calcMode="auto" fullCalcOnLoad="1"/></workbook>'));
    }
}

function addTodayDxfStyle(zip, stylesPath, styleXml) {
    const xml = strFromU8(zip[stylesPath]);
    const fillXml = `<dxf>${styleXml}</dxf>`;
    const existingDxfIndex = getExistingTodayDxfIndex(xml, fillXml);
    if (existingDxfIndex !== -1) return existingDxfIndex;
    const dxfsMatch = xml.match(/<dxfs count="(\d+)">([\s\S]*?)<\/dxfs>/);
    if (dxfsMatch) {
        const count = Number(dxfsMatch[1]);
        zip[stylesPath] = strToU8(xml.replace(dxfsMatch[0], `<dxfs count="${count + 1}">${dxfsMatch[2]}${fillXml}</dxfs>`));
        return count;
    }
    zip[stylesPath] = strToU8(xml.replace(/<dxfs[^>]*\/>/, '').replace('</styleSheet>', `<dxfs count="1">${fillXml}</dxfs></styleSheet>`));
    return 0;
}

function getExistingTodayDxfIndex(xml, fillXml) {
    const dxfsMatch = xml.match(/<dxfs count="\d+">([\s\S]*?)<\/dxfs>/);
    if (!dxfsMatch) {
        return -1;
    }

    const dxfs = dxfsMatch[1].match(/<dxf>[\s\S]*?<\/dxf>/g) || [];
    return dxfs.findIndex(dxf => dxf === fillXml);
}

function toRichTextXml(text) {
    const parts = text.split(/(\[[^\]]+\])/g).filter(part => part.length > 0);

    return parts.map(part => {
        if (part.startsWith('[') && part.endsWith(']')) {
            const comment = part.slice(1, -1);
            return [
                richTextRun('[', toArgb(colors.text)),
                richTextRun(comment, toArgb(colors.commentText)),
                richTextRun(']', toArgb(colors.text))
            ].join('');
        }

        return toSuperscriptClassXml(part, toArgb(colors.text));
    }).join('');
}

function needsRichText(text) {
    return /\[[^\]]+\]/.test(String(text)) || /\b\d+e\d+\b/.test(String(text));
}

function toSuperscriptClassXml(text, color) {
    return String(text).split(/(\b\d+e\d+\b)/g).filter(part => part.length > 0).map(part => {
        const classMatch = part.match(/^(\d+)e(\d+)$/);
        if (!classMatch) {
            return richTextRun(part, color);
        }

        const classColor = toArgb(gradeColors[classMatch[1][0]] || color);
        return [
            richTextRun(classMatch[1], classColor, false, true),
            richTextRun('e', classColor, true, true),
            richTextRun(classMatch[2], classColor, false, true)
        ].join('');
    }).join('');
}

function richTextRun(text, color, superscript = false, bold = false) {
    const verticalAlign = superscript ? '<vertAlign val="superscript"/>' : '';
    const fontWeight = bold ? '<b/>' : '';
    return `<r><rPr>${fontWeight}<color rgb="${color}"/>${verticalAlign}</rPr>${richTextValueXml(text)}</r>`;
}

function richTextValueXml(text) {
    const preserveSpace = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    return `<t${preserveSpace}>${xmlEscape(text)}</t>`;
}

function xmlEscape(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function xmlUnescape(text) {
    return String(text)
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&');
}

function toArgb(color) {
    return `FF${colorHex(color)}`;
}

function colorHex(color) {
    return String(color || '').replace(/^#/, '').toUpperCase();
}

function generateWeeklyTimetableSheet(classes) {
    const days = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi'];
    const entries = getWeeklyTimetableEntries(classes);
    const boundaries = getTimetableBoundaries(entries);
    const content = [
        [emptyCell],
        [{ v: 'Heure', s: timetableHeaderStyle }]
    ];

    days.forEach(day => {
        content[0].push({ v: day, s: timetableHeaderStyle }, { v: '', s: timetableHeaderStyle });
        content[1].push({ v: 'S1', s: timetableHeaderStyle }, { v: 'S2', s: timetableHeaderStyle });
    });

    for (let index = 0; index < boundaries.length - 1; index++) {
        const row = [{ v: formatMinutes(boundaries[index]), s: timetableHeaderStyle }];
        for (let col = 1; col <= 10; col++) {
            row.push({ v: '', s: timetableEmptyStyle });
        }
        content.push(row);
    }

    const sheet = XLSX.utils.aoa_to_sheet(content);
    sheet["!merges"] = [];

    for (let dayIndex = 0; dayIndex < days.length; dayIndex++) {
        const colIndex = 1 + dayIndex * 2;
        sheet["!merges"].push({
            s: { r: 0, c: colIndex },
            e: { r: 0, c: colIndex + 1 }
        });
    }

    entries.forEach(entry => {
        const startRow = 2 + boundaries.indexOf(entry.start);
        const endRow = 1 + boundaries.indexOf(entry.end);
        const style = getTimetableClassStyle(entry.className);
        const value = getTimetableEntryText(entry);

        if (entry.parities.size === 2) {
            const startCol = 1 + entry.dayId * 2;
            setSheetCell(sheet, startRow, startCol, { v: value, s: style });
            mergeIfNeeded(sheet, startRow, startCol, endRow, startCol + 1);
        } else {
            const parity = Array.from(entry.parities)[0];
            const startCol = 1 + entry.dayId * 2 + parity - 1;
            setSheetCell(sheet, startRow, startCol, { v: value, s: style });
            mergeIfNeeded(sheet, startRow, startCol, endRow, startCol);
        }
    });

    sheet["!cols"] = [
        { wch: 8 },
        { wch: 18 }, { wch: 18 },
        { wch: 18 }, { wch: 18 },
        { wch: 18 }, { wch: 18 },
        { wch: 18 }, { wch: 18 },
        { wch: 18 }, { wch: 18 }
    ];
    sheet["!rows"] = [
        { hpt: 22 },
        { hpt: 18 },
        ...Array(Math.max(0, boundaries.length - 1)).fill({ hpt: 58 })
    ];

    return sheet;
}

function getWeeklyTimetableEntries(classes) {
    const entriesByKey = new Map();
    const configuredStarts = getConfiguredTimetableSlots().map(slot => slot.start);

    classes.forEach(c => {
        for (let week = weeks.first; week <= weeks.last; week++) {
            if (weeks.holidays.includes(week)) {
                continue;
            }

            for (let dayOffset = 0; dayOffset < weeks.length; dayOffset++) {
                const day = getDay(week, dayOffset, weeks.startDate);
                if (day.id > 4 || weeks.formattedOffDays.includes(day.formattedDate)) {
                    continue;
                }

                getShownTimes(c, day).forEach(time => {
                    const start = getTimetableEntryStart(timeToMinutes(time.time), configuredStarts);
                    const end = start + getTimetableEntryDurationMinutes(time);
                    const key = [
                        day.id,
                        start,
                        end,
                        c.name,
                        time.comment || '',
                        time.subject || '',
                        time.location || ''
                    ].join('|');

                    if (!entriesByKey.has(key)) {
                        entriesByKey.set(key, {
                            dayId: day.id,
                            start,
                            end,
                            className: c.name,
                            comment: time.comment,
                            subject: time.subject,
                            location: time.location,
                            parities: new Set()
                        });
                    }

                    entriesByKey.get(key).parities.add(week % 2 === 1 ? 1 : 2);
                });
            }
        }
    });

    return Array.from(entriesByKey.values())
        .sort((a, b) => a.dayId - b.dayId || a.start - b.start || a.className.localeCompare(b.className, 'fr', { numeric: true }));
}

function getTimetableEntryStart(start, configuredStarts) {
    if (!configuredStarts.length) {
        return start;
    }

    return configuredStarts.reduce((closest, candidate) => {
        return Math.abs(candidate - start) < Math.abs(closest - start) ? candidate : closest;
    }, configuredStarts[0]);
}

function getTimetableEntryDurationMinutes(time) {
    return Math.round(getLessonDuration(time) * 60);
}

function getTimetableBoundaries(entries) {
    const configuredSlots = getConfiguredTimetableSlots();
    if (configuredSlots.length) {
        return Array.from(new Set(trimEmptyEdgeSlots(configuredSlots, entries).flatMap(slot => [slot.start, slot.end]))).sort((a, b) => a - b);
    }

    const boundaries = new Set();
    entries.forEach(entry => {
        boundaries.add(entry.start);
        boundaries.add(entry.end);
    });

    if (boundaries.size === 0) {
        boundaries.add(8 * 60);
        boundaries.add(17 * 60);
    }

    return Array.from(boundaries).sort((a, b) => a - b);
}

function getConfiguredTimetableSlots() {
    if (!schoolHours || !Array.isArray(schoolHours.starts) || schoolHours.starts.length === 0) {
        return [];
    }

    const duration = getConfiguredLessonDurationMinutes();
    return schoolHours.starts
        .map(start => timeToMinutes(start))
        .filter(start => Number.isFinite(start))
        .sort((a, b) => a - b)
        .map(start => ({ start, end: start + duration }));
}

function trimEmptyEdgeSlots(slots, entries) {
    if (!schoolHours.hideEmptyEdges || slots.length === 0) {
        return slots;
    }

    let first = 0;
    let last = slots.length - 1;

    while (first <= last && !hasEntryInSlot(slots[first], entries)) {
        first++;
    }

    while (last >= first && !hasEntryInSlot(slots[last], entries)) {
        last--;
    }

    return first <= last ? slots.slice(first, last + 1) : slots;
}

function hasEntryInSlot(slot, entries) {
    return entries.some(entry => entry.start < slot.end && entry.end > slot.start);
}

function getConfiguredLessonDurationMinutes() {
    return Number.isFinite(Number(schoolHours.lessonDurationMinutes))
        ? Number(schoolHours.lessonDurationMinutes)
        : 55;
}

function timeToMinutes(time) {
    const decimalTime = String(time);
    const separator = decimalTime.includes(':') ? ':' : '.';
    const [hours, minutes = '0'] = decimalTime.split(separator);

    return Number(hours) * 60 + Number(minutes.padEnd(2, '0'));
}

function formatMinutes(minutes) {
    const hours = Math.floor(minutes / 60);
    const mins = minutes % 60;

    return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
}

function getTimetableClassStyle(className) {
    const defaultGradeColor = gradeColors[className[Number(0)]] || dateColor;
    const gradeColor = fun.blendColors(white, defaultGradeColor, 0.75);

    return {
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        font: { color: { rgb: colorHex(colors.text) }, sz: 10 },
        fill: { fgColor: { rgb: fun.blendColors(white, gradeColor, 0.5) } },
        border: {
            top: borderStyle,
            bottom: borderStyle,
            left: borderStyle,
            right: borderStyle
        }
    };
}

function getTimetableEntryText(entry) {
    const subject = formatSubject(entry.subject);
    const comment = formatWeeklyComment(entry.comment, entry.subject);
    const lines = [];

    if (shouldShowWeeklySubject() && subject) {
        lines.push(subject);
    }
    lines.push(formatClassWithComment(entry.className, comment));
    if (shouldShowWeeklyRoom() && entry.location) {
        lines.push(entry.location);
    }

    return lines.join('\n');
}

function formatSubject(subject) {
    const value = String(subject || 'MATHEMATIQUES').trim();

    return value.replace(/\bVIE DE CLASSE\b/gi, 'VDC');
}

function shouldShowWeeklyRoom() {
    return !config.weeklyTimetable || config.weeklyTimetable.showRoom !== false;
}

function shouldShowWeeklySubject() {
    return !!(config.weeklyTimetable && config.weeklyTimetable.showSubject);
}

function getConfigNumber(value, fallback) {
    return Number.isFinite(Number(value)) ? Number(value) : fallback;
}

function getWeeklyTimetablePdfLayout() {
    const pdfConfig = (config.weeklyTimetable && config.weeklyTimetable.pdf) || {};
    const fontConfig = pdfConfig.fonts || {};
    const paddingConfig = pdfConfig.lessonPadding || {};

    return {
        margin: getConfigNumber(pdfConfig.margin, 18),
        timeColWidth: getConfigNumber(pdfConfig.timeColWidth, 44),
        headerTop: getConfigNumber(pdfConfig.headerTop, 42),
        dayHeaderHeight: getConfigNumber(pdfConfig.dayHeaderHeight, 18),
        weekHeaderHeight: getConfigNumber(pdfConfig.weekHeaderHeight, 16),
        rowHeightScale: getConfigNumber(pdfConfig.rowHeightScale, 1),
        breakRowHeightScale: getConfigNumber(pdfConfig.breakRowHeightScale, 1),
        lineWidth: getConfigNumber(pdfConfig.lineWidth, 1),
        emptyLineWidth: getConfigNumber(pdfConfig.emptyLineWidth, 1),
        fonts: {
            title: getConfigNumber(fontConfig.title, 15),
            dayHeader: getConfigNumber(fontConfig.dayHeader, 10),
            weekHeader: getConfigNumber(fontConfig.weekHeader, 8),
            time: getConfigNumber(fontConfig.time, 8),
            lesson: getConfigNumber(fontConfig.lesson, 7.5),
            lessonLineHeight: getConfigNumber(fontConfig.lessonLineHeight, 9)
        },
        lessonPadding: {
            x: getConfigNumber(paddingConfig.x, 4),
            y: getConfigNumber(paddingConfig.y, 6)
        }
    };
}

function setSheetCell(sheet, row, col, cell) {
    const address = XLSX.utils.encode_cell({ r: row, c: col });
    sheet[address] = { t: 's', v: cell.v, s: cell.s };
}

function mergeIfNeeded(sheet, startRow, startCol, endRow, endCol) {
    if (startRow === endRow && startCol === endCol) {
        return;
    }

    sheet["!merges"].push({
        s: { r: startRow, c: startCol },
        e: { r: endRow, c: endCol }
    });
}

function generateWeeklyTimetablePdf(classes, outputPath, orientation = 'landscape') {
    const entries = getWeeklyTimetableEntries(classes);
    const slots = getPdfTimetableSlots(entries);
    const layout = getWeeklyTimetablePdfLayout();
    const page = getPdfPage(orientation);
    currentPdfPageHeight = page.height;
    const margin = layout.margin;
    const timeColWidth = layout.timeColWidth;
    const headerTop = layout.headerTop;
    const dayHeaderHeight = layout.dayHeaderHeight;
    const weekHeaderHeight = layout.weekHeaderHeight;
    const gridTop = headerTop + dayHeaderHeight + weekHeaderHeight;
    const gridWidth = page.width - margin * 2 - timeColWidth;
    const availableGridHeight = page.height - margin - gridTop;
    const gridHeight = availableGridHeight * layout.rowHeightScale;
    const dayWidth = gridWidth / 5;
    const halfDayWidth = dayWidth / 2;
    const slotMetrics = getPdfSlotMetrics(slots, gridTop, gridHeight, layout.breakRowHeightScale);
    const commands = [];

    drawRect(commands, 0, 0, page.width, page.height, colorHex(colors.white));
    drawText(commands, getWeeklyTimetableTitle(classes), margin, 14, layout.fonts.title, colorHex(colors.text), 'bold');
    setLineWidth(commands, layout.lineWidth);

    ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi'].forEach((day, dayIndex) => {
        const x = margin + timeColWidth + dayIndex * dayWidth;
        drawRect(commands, x, headerTop, dayWidth, dayHeaderHeight, dateColor, colorHex(colors.text));
        drawTextCentered(commands, day, x, headerTop + 4, dayWidth, layout.fonts.dayHeader, colorHex(colors.text), 'bold');
        drawRect(commands, x, headerTop + dayHeaderHeight, halfDayWidth, weekHeaderHeight, dateColor, colorHex(colors.text));
        drawRect(commands, x + halfDayWidth, headerTop + dayHeaderHeight, halfDayWidth, weekHeaderHeight, dateColor, colorHex(colors.text));
        drawTextCentered(commands, 'S1', x, headerTop + dayHeaderHeight + 4, halfDayWidth, layout.fonts.weekHeader, colorHex(colors.text), 'normal');
        drawTextCentered(commands, 'S2', x + halfDayWidth, headerTop + dayHeaderHeight + 4, halfDayWidth, layout.fonts.weekHeader, colorHex(colors.text), 'normal');
    });

    for (let index = 0; index < slots.length; index++) {
        const slot = slots[index];
        const { y, height: rowHeight } = slotMetrics[index];
        drawRect(commands, margin, y, timeColWidth, rowHeight, colorHex(colors.white), colorHex(colors.text));
        drawTextCentered(commands, formatMinutes(slot.start), margin, y + 4, timeColWidth, layout.fonts.time, colorHex(colors.text), 'normal');
        drawTextCentered(commands, formatMinutes(slot.end), margin, y + rowHeight - layout.fonts.time - 4, timeColWidth, layout.fonts.time, colorHex(colors.text), 'normal');

        for (let dayIndex = 0; dayIndex < 5; dayIndex++) {
            const x = margin + timeColWidth + dayIndex * dayWidth;
            const emptyFill = slot.break ? colors.timetableBreak : colors.timetableEmpty;
            if (shouldSplitEmptyPdfCell(entries, slot, dayIndex)) {
                drawRect(commands, x, y, halfDayWidth, rowHeight, colorHex(emptyFill), colorHex(colors.timetableEmptyBorder));
                drawRect(commands, x + halfDayWidth, y, halfDayWidth, rowHeight, colorHex(emptyFill), colorHex(colors.timetableEmptyBorder));
            } else {
                drawRect(commands, x, y, dayWidth, rowHeight, colorHex(emptyFill), colorHex(colors.timetableEmptyBorder));
            }
            setLineWidth(commands, layout.emptyLineWidth);
            drawLine(commands, x + 20, y + rowHeight / 2, x + dayWidth - 20, y + rowHeight / 2, colorHex(colors.timetableEmptyLine));
            setLineWidth(commands, layout.lineWidth);
        }
    }

    entries.forEach(entry => {
        const startIndex = slots.findIndex(slot => slot.start === entry.start);
        if (startIndex === -1) {
            return;
        }
        const endIndex = getPdfEntryEndSlotIndex(slots, entry, startIndex);
        const y = slotMetrics[startIndex].y;
        const h = getPdfEntryHeight(slotMetrics, startIndex, endIndex);
        const baseX = margin + timeColWidth + entry.dayId * dayWidth;
        const isFullDay = entry.parities.size === 2;
        const parity = Array.from(entry.parities)[0];
        const x = isFullDay ? baseX : baseX + (parity - 1) * halfDayWidth;
        const w = isFullDay ? dayWidth : halfDayWidth;
        const color = getPdfClassColor(entry.className);
        const lines = getTimetableEntryLines(entry);

        drawRect(commands, x, y, w, h, color.fill, color.border);
        drawWrappedCenteredLines(commands, lines, x + layout.lessonPadding.x, y + layout.lessonPadding.y, w - layout.lessonPadding.x * 2, h - layout.lessonPadding.y * 2, colorHex(colors.text), layout.fonts.lesson, layout.fonts.lessonLineHeight);
    });

    fs.writeFileSync(outputPath, buildPdf(page.width, page.height, commands.join('\n')));
}

function getPdfPage(orientation) {
    return orientation === 'portrait'
        ? { width: 595, height: 842 }
        : { width: 842, height: 595 };
}

function getTimetableEntryLines(entry) {
    const subject = formatSubject(entry.subject);
    const comment = formatWeeklyComment(entry.comment, entry.subject);
    const lines = [];

    if (shouldShowWeeklySubject() && subject) {
        lines.push(subject);
    }
    lines.push(formatClassWithComment(entry.className, comment));
    if (shouldShowWeeklyRoom() && entry.location) {
        lines.push(entry.location);
    }

    return lines;
}

function formatClassWithComment(className, comment) {
    return [formatClassName(className), comment].filter(Boolean).join(' ');
}

function getWeeklyTimetableTitle(classes) {
    return getNamedTitle(classes, 'EDT');
}

function getPlanningTitle(classes) {
    return getNamedTitle(classes, 'Planning');
}

function getNamedTitle(classes, prefix) {
    const metadata = classes.metadata || {};
    const parts = [prefix];

    if (metadata.school) {
        parts.push(metadata.school);
    }
    if (metadata.teacher) {
        parts.push(metadata.teacher);
    }
    if (metadata.years) {
        parts.push(metadata.years);
    }

    return parts.join(' - ');
}

function getOutputPath(classes, prefix, extension) {
    return path.join(path.dirname(config.output), `${toKebabCase(getNamedTitle(classes, prefix))}.${extension}`);
}

function getPdfOrientations() {
    const pdfConfig = (config.weeklyTimetable && config.weeklyTimetable.pdf) || {};
    if (Array.isArray(pdfConfig.orientations) && pdfConfig.orientations.length) {
        return pdfConfig.orientations;
    }

    return [pdfConfig.orientation || 'landscape'];
}

function getPdfOutputPath(classes, orientation) {
    const baseName = toKebabCase(getWeeklyTimetableTitle(classes));
    const suffix = orientation === 'portrait' ? '-portrait' : '';

    return path.join(path.dirname(config.output), `${baseName}${suffix}.pdf`);
}

function getPdfSlotMetrics(slots, gridTop, gridHeight, breakRowHeightScale) {
    const weights = slots.map(slot => slot.break ? breakRowHeightScale : 1);
    const totalWeight = weights.reduce((sum, weight) => sum + weight, 0) || 1;
    const unitHeight = gridHeight / totalWeight;
    let y = gridTop;

    return slots.map((slot, index) => {
        const height = unitHeight * weights[index];
        const metric = { y, height };
        y += height;
        return metric;
    });
}

function getPdfEntryHeight(slotMetrics, startIndex, endIndex) {
    return slotMetrics
        .slice(startIndex, endIndex)
        .reduce((height, metric) => height + metric.height, 0);
}

function shouldSplitEmptyPdfCell(entries, slot, dayId) {
    return entries.some(entry => {
        return entry.dayId === dayId
            && entry.parities.size === 1
            && entry.start < slot.end
            && entry.end > slot.start;
    });
}

function getPdfTimetableSlots(entries) {
    const configuredSlots = getConfiguredTimetableSlots();
    if (configuredSlots.length) {
        return getPdfSlotsFromConfiguredSlots(trimEmptyEdgeSlots(configuredSlots, entries));
    }

    const starts = Array.from(new Set(entries.map(entry => entry.start))).sort((a, b) => a - b);
    if (starts.length === 0) {
        return [{ start: 8 * 60, end: 17 * 60 }];
    }

    const slots = [];
    let previousEnd = null;

    starts.forEach(start => {
        if (previousEnd != null && start - previousEnd >= 30) {
            slots.push({ start: previousEnd, end: start, break: true });
        }

        const end = Math.max(...entries
            .filter(entry => entry.start === start)
            .map(entry => entry.end));
        slots.push({ start, end });
        previousEnd = Math.max(previousEnd || end, end);
    });

    return slots;
}

function getPdfSlotsFromConfiguredSlots(configuredSlots) {
    const slots = [];
    let previousEnd = null;

    configuredSlots.forEach(slot => {
        if (previousEnd != null && slot.start - previousEnd >= 30) {
            slots.push({ start: previousEnd, end: slot.start, break: true });
        }

        slots.push(slot);
        previousEnd = slot.end;
    });

    return slots;
}

function getPdfEntryEndSlotIndex(slots, entry, startIndex) {
    const endIndex = slots.findIndex((slot, index) => index > startIndex && slot.start >= entry.end);

    return endIndex === -1 ? startIndex + 1 : endIndex;
}

function toKebabCase(value) {
    return String(value)
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-zA-Z0-9]+/g, '-')
        .replace(/^-+|-+$/g, '')
        .toLowerCase();
}

function formatWeeklyComment(comment, subject) {
    const subjectComment = subject && String(comment).trim().toUpperCase() === String(subject).trim().toUpperCase();
    if (subjectComment && !shouldShowWeeklySubject()) {
        return isVieDeClasse(subject) ? '[VDC]' : '';
    }

    if (!comment) {
        return isVieDeClasse(subject) && !shouldShowWeeklySubject() ? '[VDC]' : '';
    }

    return formatComment(comment);
}

function isVieDeClasse(value) {
    return /\bVIE DE CLASSE\b/i.test(String(value || ''));
}

function getPdfClassColor(className) {
    const defaultGradeColor = gradeColors[className[Number(0)]] || dateColor;
    const border = fun.blendColors(white, defaultGradeColor, 0.85);
    const fill = fun.blendColors(white, border, 0.35);

    return { fill, border };
}

function buildPdf(width, height, stream) {
    const objects = [
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${width} ${height}] /Resources << /Font << /F1 4 0 R /F2 5 0 R >> >> /Contents 6 0 R >>`,
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>',
        `<< /Length ${Buffer.byteLength(stream, 'utf8')} >>\nstream\n${stream}\nendstream`
    ];
    const parts = ['%PDF-1.4\n'];
    const offsets = [0];

    objects.forEach((object, index) => {
        offsets.push(Buffer.byteLength(parts.join(''), 'utf8'));
        parts.push(`${index + 1} 0 obj\n${object}\nendobj\n`);
    });

    const xrefOffset = Buffer.byteLength(parts.join(''), 'utf8');
    parts.push(`xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`);
    offsets.slice(1).forEach(offset => {
        parts.push(`${String(offset).padStart(10, '0')} 00000 n \n`);
    });
    parts.push(`trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`);

    return Buffer.from(parts.join(''), 'utf8');
}

function drawRect(commands, x, y, width, height, fill, stroke = null) {
    const [r, g, b] = hexToRgb(fill);
    commands.push(`${r} ${g} ${b} rg`);
    if (stroke) {
        const [sr, sg, sb] = hexToRgb(stroke);
        commands.push(`${sr} ${sg} ${sb} RG`);
    }
    commands.push(`${formatPdfNumber(x)} ${formatPdfNumber(pdfY(y, height))} ${formatPdfNumber(width)} ${formatPdfNumber(height)} re ${stroke ? 'B' : 'f'}`);
}

function drawLine(commands, x1, y1, x2, y2, stroke) {
    const [r, g, b] = hexToRgb(stroke);
    commands.push(`${r} ${g} ${b} RG`);
    commands.push(`${formatPdfNumber(x1)} ${formatPdfNumber(pdfY(y1, 0))} m ${formatPdfNumber(x2)} ${formatPdfNumber(pdfY(y2, 0))} l S`);
}

function setLineWidth(commands, width) {
    commands.push(`${formatPdfNumber(width)} w`);
}

function drawText(commands, text, x, y, size, color, weight = 'normal') {
    const [r, g, b] = hexToRgb(color);
    const font = weight === 'bold' ? 'F2' : 'F1';
    commands.push(`BT /${font} ${size} Tf ${r} ${g} ${b} rg ${formatPdfNumber(x)} ${formatPdfNumber(pdfY(y, size))} Td (${escapePdfText(text)}) Tj ET`);
}

function drawTextCentered(commands, text, x, y, width, size, color, weight = 'normal') {
    const textWidth = estimateTextWidth(text, size);
    drawText(commands, text, x + Math.max(0, (width - textWidth) / 2), y, size, color, weight);
}

function drawWrappedCenteredLines(commands, lines, x, y, width, height, color, size = 7.5, lineHeight = 9) {
    const wrapped = [];
    lines.forEach(line => wrapped.push(...wrapText(line, width, size)));
    const startY = y + Math.max(0, (height - wrapped.length * lineHeight) / 2);

    wrapped.forEach((line, index) => {
        drawRichTextCentered(commands, line, x, startY + index * lineHeight, width, size, color);
    });
}

function drawRichTextCentered(commands, text, x, y, width, size, color) {
    const parts = splitRichPdfText(text, color, colorHex(colors.commentText), size);
    const textWidth = parts.reduce((total, part) => total + estimateTextWidth(part.text, part.size), 0);
    let currentX = x + Math.max(0, (width - textWidth) / 2);

    parts.forEach(part => {
        drawText(commands, part.text, currentX, y + part.yOffset, part.size, part.color, part.bold ? 'bold' : 'normal');
        currentX += estimateTextWidth(part.text, part.size);
    });
}

function splitRichPdfText(text, color, bracketTextColor, size) {
    const parts = [];
    String(text).split(/(\[[^\]]+\])/g).filter(Boolean).forEach(part => {
        const bracketMatch = part.match(/^\[([^\]]+)\]$/);
        if (!bracketMatch) {
            parts.push(...splitPdfClassSuperscript(part, color, size));
            return;
        }

        parts.push({ text: '[', color, size, yOffset: 0 });
        parts.push({ text: bracketMatch[1], color: bracketTextColor, size, yOffset: 0 });
        parts.push({ text: ']', color, size, yOffset: 0 });
    });

    return parts;
}

function splitPdfClassSuperscript(text, color, size) {
    return String(text).split(/(\b\d+e\d+\b)/g).filter(Boolean).flatMap(part => {
        const classMatch = part.match(/^(\d+)e(\d+)$/);
        if (!classMatch) {
            return [{ text: part, color, size, yOffset: 0, bold: false }];
        }

        const classColor = gradeColors[classMatch[1][0]] || color;
        return [
            { text: classMatch[1], color: classColor, size, yOffset: 0, bold: true },
            { text: 'e', color: classColor, size: size * 0.72, yOffset: -size * 0.25, bold: true },
            { text: classMatch[2], color: classColor, size, yOffset: 0, bold: true }
        ];
    });
}

function wrapText(text, width, size) {
    const words = String(text).split(/\s+/);
    const lines = [];
    let line = '';

    words.forEach(word => {
        const candidate = line ? `${line} ${word}` : word;
        if (estimateTextWidth(candidate, size) <= width || !line) {
            line = candidate;
            return;
        }

        lines.push(line);
        line = word;
    });

    if (line) {
        lines.push(line);
    }

    return lines;
}

function estimateTextWidth(text, size) {
    return Array.from(String(text)).reduce((width, char) => {
        return width + getEstimatedCharWidth(char) * size;
    }, 0);
}

function getEstimatedCharWidth(char) {
    if (char === ' ') {
        return 0.28;
    }
    if (/[\[\]\(\)]/.test(char)) {
        return 0.3;
    }
    if (/[A-Z]/.test(char)) {
        return 0.68;
    }
    if (/[0-9]/.test(char)) {
        return 0.56;
    }
    if (/[il.,:;]/.test(char)) {
        return 0.25;
    }
    if (/[mw]/.test(char)) {
        return 0.75;
    }

    return 0.5;
}

function hexToRgb(hex) {
    const value = hex.replace('#', '');
    return [
        Number.parseInt(value.slice(0, 2), 16) / 255,
        Number.parseInt(value.slice(2, 4), 16) / 255,
        Number.parseInt(value.slice(4, 6), 16) / 255
    ].map(formatPdfNumber);
}

function pdfY(y, height) {
    return currentPdfPageHeight - y - height;
}

function formatPdfNumber(value) {
    return Number(value).toFixed(3).replace(/\.?0+$/, '');
}

function escapePdfText(text) {
    return String(text)
        .replace(/[^\x20-\x7E]/g, '')
        .replace(/\\/g, '\\\\')
        .replace(/\(/g, '\\(')
        .replace(/\)/g, '\\)');
}

function writeSheet() {
    const classes = fun.loadClassesSchedules();
    if (classes.weeks) {
        Object.assign(weeks, classes.weeks);
    }

    const outputPath = getOutputPath(classes, 'planning', config.tableType);
    const legacyOutputPath = config.output + "." + config.tableType;
    previousComments = loadPreviousComments(classes, fs.existsSync(outputPath) ? outputPath : legacyOutputPath);

    const workbook = XLSX.utils.book_new();
    const weeksSheet = generateWeeksSheet(classes);
    XLSX.utils.book_append_sheet(workbook, weeksSheet, weeksSheetName);
    XLSX.writeFile(workbook, outputPath, { bookType: config.tableType });
    colorizeCommentCells(outputPath, classes);
    console.log(`Sheet written at : ${outputPath}`);
    getPdfOrientations().forEach(orientation => {
        const pdfOutputPath = getPdfOutputPath(classes, orientation);
        generateWeeklyTimetablePdf(classes, pdfOutputPath, orientation);
        console.log(`Weekly timetable PDF written at : ${pdfOutputPath}`);
    });
}

module.exports = () => {
    writeSheet();
};
