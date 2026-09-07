const fs = require('fs');
const path = require('path');
const offDayNames = require(path.join(__basedir, 'config/off-days'));
const schoolHours = require(path.join(__basedir, 'config/school-hours'));

const dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];

module.exports = () => {
    const icsPath = findIcalFile();
    if (!icsPath) {
        console.warn('icalExport is true, but no .ics file was found in input/. Falling back to .js schedules.');
        return null;
    }

    return loadIcalSchedules(icsPath);
};

function findIcalFile() {
    const inputDir = path.join(__basedir, config.input);
    const files = fs.readdirSync(inputDir);
    const icsFile = files.find(file => file.endsWith('.ics'));

    return icsFile ? path.join(inputDir, icsFile) : null;
}

function loadIcalSchedules(icsPath) {
    const content = fs.readFileSync(icsPath, 'utf8');
    const events = parseEvents(content);
    const classesByName = new Map();
    const offDays = [];
    const offDayLabels = {};
    let firstDate = null;
    let lastDate = null;
    let teacher = null;
    let vacationIndex = 0;
    let holidayIndex = 0;

    events.forEach(event => {
        if (isHolidayEvent(event)) {
            const dates = getDateRange(parseIcalDate(event.DTSTART), parseIcalDate(event.DTEND));
            const label = getOffDayLabel(event, vacationIndex, holidayIndex);
            offDays.push(...dates);
            dates.forEach(date => {
                offDayLabels[formatDateKey(date)] = label;
            });
            if (isVacationSummary(event.SUMMARY)) {
                vacationIndex++;
            } else if (isHolidaySummary(event.SUMMARY)) {
                holidayIndex++;
            }
            return;
        }

        if (!isCourseEvent(event) || isCancelledEvent(event)) {
            return;
        }

        const startDate = parseIcalDate(event.DTSTART);
        const endDate = parseIcalDate(event.DTEND);
        if (!startDate) {
            return;
        }

        const classInfo = getClassInfo(event);
        if (!classInfo.name) {
            return;
        }
        teacher = teacher || classInfo.teacher;

        firstDate = minDate(firstDate, startDate);
        lastDate = maxDate(lastDate, startDate);

        if (!classesByName.has(classInfo.name)) {
            classesByName.set(classInfo.name, { name: classInfo.name });
        }

        const schedule = classesByName.get(classInfo.name);
        const dayName = dayNames[startDate.getDay()];

        if (!schedule[dayName]) {
            schedule[dayName] = { class: [] };
        }

        schedule[dayName].class.push({
            time: dateToDecimalTime(startDate),
            date: formatDateKey(startDate),
            duration: getDurationHours(startDate, endDate),
            subject: classInfo.subject,
            location: event.LOCATION || '',
            comment: classInfo.comment
        });
    });

    const classes = Array.from(classesByName.values())
        .sort((a, b) => a.name.localeCompare(b.name, 'fr', { numeric: true }));

    classes.forEach(schedule => {
        dayNames.forEach(dayName => {
            if (schedule[dayName]) {
                schedule[dayName].class.sort((a, b) => a.time - b.time);
            }
        });
    });

    classes.weeks = buildWeeksConfig(content, firstDate, lastDate, offDays, offDayLabels);
    classes.metadata = buildMetadata(content, classes.weeks, teacher);

    return classes;
}

function parseEvents(content) {
    const unfolded = content.replace(/\r?\n[ \t]/g, '');
    const blocks = unfolded.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g) || [];

    return blocks.map(block => {
        const event = {};
        block.split(/\r?\n/).forEach(line => {
            const separatorIndex = line.indexOf(':');
            if (separatorIndex === -1) {
                return;
            }

            const key = line.slice(0, separatorIndex).split(';')[0];
            const value = unescapeIcalText(line.slice(separatorIndex + 1));
            event[key] = value;
        });
        return event;
    });
}

function isHolidayEvent(event) {
    return event.CATEGORIES && event.CATEGORIES.includes('Jours');
}

function getOffDayLabel(event, vacationIndex, holidayIndex) {
    if (isVacationSummary(event.SUMMARY)) {
        return offDayNames.vacations[vacationIndex] || event.SUMMARY || 'Vacances';
    }

    if (isHolidaySummary(event.SUMMARY)) {
        return offDayNames.holidays[holidayIndex] || event.SUMMARY || 'Jour férié';
    }

    return event.SUMMARY || 'Jour férié';
}

function isVacationSummary(summary) {
    return /vacances/i.test(String(summary || ''));
}

function isHolidaySummary(summary) {
    return /férié|ferie/i.test(String(summary || ''));
}

function isCourseEvent(event) {
    return event.CATEGORIES && event.CATEGORIES.includes('Cours');
}

function isCancelledEvent(event) {
    return (event.CATEGORIES && event.CATEGORIES.includes('annul'))
        || (event.SUMMARY && event.SUMMARY.includes('Cours annul'));
}

function getClassInfo(event) {
    const description = event.DESCRIPTION || '';
    const className = matchFirst(description, [
        /Classe\s*:\s*([^\n]+)/,
        /Partie de classe\s*:\s*(?:<([^>]+)>|&lt;([^&]+)&gt;)/
    ]);

    if (!className) {
        return { name: null, comment: null, subject: null, teacher: null };
    }

    const group = matchFirst(description, [
        /Groupe\s*:\s*\[([^\]]+)\]/,
        /Partie de classe\s*:\s*(?:<[^>]+>|&lt;[^&]+&gt;)\s*([^\n]+)/
    ]);

    const subject = matchFirst(description, [/Matière\s*:\s*([^\n]+)/]);
    const teacher = matchFirst(description, [/Professeur\s*:\s*([^\n]+)/]);
    const commentParts = [];

    if (group && group.trim() !== className.trim()) {
        commentParts.push(group.trim());
    }
    if (subject && subject.trim() !== 'MATHEMATIQUES') {
        commentParts.push(subject.trim());
    }

    return {
        name: className.trim(),
        subject: subject ? subject.trim() : null,
        teacher: teacher ? teacher.trim() : null,
        comment: commentParts.length ? commentParts.join(' - ') : null
    };
}

function matchFirst(text, patterns) {
    for (const pattern of patterns) {
        const match = text.match(pattern);
        if (match) {
            return (match[1] || match[2] || '').trim();
        }
    }

    return null;
}

function parseIcalDate(value) {
    if (!value) {
        return null;
    }

    const normalized = value.trim();
    const dateMatch = normalized.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (dateMatch) {
        return new Date(Number(dateMatch[1]), Number(dateMatch[2]) - 1, Number(dateMatch[3]));
    }

    const dateTimeMatch = normalized.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z?$/);
    if (!dateTimeMatch) {
        return null;
    }

    const [, year, month, day, hour, minute, second] = dateTimeMatch;
    if (normalized.endsWith('Z')) {
        return new Date(Date.UTC(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second)));
    }

    return new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second));
}

function dateToDecimalTime(date) {
    return Number(`${date.getHours()}.${String(date.getMinutes()).padStart(2, '0')}`);
}

function getDurationHours(startDate, endDate) {
    if (schoolHours.lessonDurationMinutes) {
        return schoolHours.lessonDurationMinutes / 60;
    }

    if (config.lessonDurationMinutes) {
        return config.lessonDurationMinutes / 60;
    }

    if (!endDate) {
        return 1;
    }

    return Math.max(0, (endDate - startDate) / (60 * 60 * 1000));
}

function formatDateKey(date) {
    return date.toLocaleDateString('fr-FR');
}

function getDateRange(startDate, endDate) {
    const dates = [];
    if (!startDate || !endDate) {
        return dates;
    }

    const currentDate = new Date(startDate);
    const lastDate = new Date(endDate);

    while (currentDate < lastDate) {
        dates.push(new Date(currentDate));
        currentDate.setDate(currentDate.getDate() + 1);
    }

    return dates;
}

function buildWeeksConfig(content, firstDate, lastDate, offDays, offDayLabels = {}) {
    const calendarStart = parseCalendarBoundary(content, 'X-CALSTART') || firstDate;
    const calendarEnd = parseCalendarBoundary(content, 'X-CALEND') || lastDate;
    const startDate = getMonday(calendarStart || new Date());
    const endDate = calendarEnd || lastDate || startDate;
    const lastWeek = Math.max(1, Math.ceil((daysBetween(startDate, endDate) + 1) / 7));

    return {
        first: 1,
        last: lastWeek,
        holidays: getFullHolidayWeeks(startDate, lastWeek, offDays),
        startDate,
        offDays,
        offDayLabels,
        length: 5
    };
}

function buildMetadata(content, weeksConfig, teacher) {
    return {
        school: getSchoolName(content),
        teacher: formatTeacherName(teacher || getCalendarTeacherName(content)),
        years: getSchoolYears(weeksConfig.startDate)
    };
}

function getSchoolName(content) {
    const unfolded = content.replace(/\r?\n[ \t]/g, '');
    const description = matchFirst(unfolded, [/X-WR-CALDESC[^:]*:([^\n]+)/]);
    const school = description && description.match(/\s-\s(COLLEGE[^-]+)/i);

    return school ? titleCase(school[1].trim()) : '';
}

function getCalendarTeacherName(content) {
    const name = matchFirst(content, [/X-WR-CALNAME[^:]*:Calendrier\s*-\s*([^-]+)/]);
    if (!name) {
        return null;
    }

    return name.trim().split(/\s+/)[0];
}

function formatTeacherName(teacher) {
    if (!teacher) {
        return '';
    }

    const lastName = teacher.trim().split(/\s+/)[0];
    return `M.${titleCase(lastName)}`;
}

function getSchoolYears(startDate) {
    const startYear = startDate.getFullYear();
    return `${startYear}-${startYear + 1}`;
}

function titleCase(value) {
    return String(value).toLowerCase().replace(/(^|\s|-)(\S)/g, (match, separator, letter) => `${separator}${letter.toUpperCase()}`);
}

function parseCalendarBoundary(content, property) {
    const match = content.match(new RegExp(`${property}:(\\d{8})`));
    return match ? parseIcalDate(match[1]) : null;
}

function getMonday(date) {
    const monday = new Date(date);
    const day = monday.getDay();
    const offset = day === 0 ? -6 : 1 - day;
    monday.setDate(monday.getDate() + offset);
    monday.setHours(0, 0, 0, 0);
    return monday;
}

function getFullHolidayWeeks(startDate, lastWeek, offDays) {
    const offDayKeys = new Set(offDays.map(formatDateKey));
    const holidays = [];

    for (let week = 1; week <= lastWeek; week++) {
        let fullHolidayWeek = true;
        for (let dayOffset = 0; dayOffset < 5; dayOffset++) {
            const date = new Date(startDate);
            date.setDate(startDate.getDate() + (week - 1) * 7 + dayOffset);
            fullHolidayWeek = fullHolidayWeek && offDayKeys.has(formatDateKey(date));
        }

        if (fullHolidayWeek) {
            holidays.push(week);
        }
    }

    return holidays;
}

function daysBetween(startDate, endDate) {
    const start = new Date(startDate);
    const end = new Date(endDate);
    start.setHours(0, 0, 0, 0);
    end.setHours(0, 0, 0, 0);

    return Math.floor((end - start) / (24 * 60 * 60 * 1000));
}

function minDate(current, candidate) {
    return !current || candidate < current ? candidate : current;
}

function maxDate(current, candidate) {
    return !current || candidate > current ? candidate : current;
}

function unescapeIcalText(value) {
    return value
        .replace(/\\n/g, '\n')
        .replace(/\\,/g, ',')
        .replace(/\\;/g, ';')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>');
}
