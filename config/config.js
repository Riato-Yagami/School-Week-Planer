module.exports = {
    input : 'input/',
    output : 'output/week-planer',
    tableType : 'xlsx',
    icalExport : true,
    weeklyTimetable : {
        showRoom : false,
        showSubject : false,
        pdf : {
            orientations : ['landscape', 'portrait'],
            margin : 18,
            timeColWidth : 44,
            headerTop : 42,
            dayHeaderHeight : 20,
            weekHeaderHeight : 16,
            rowHeightScale : 0.94,
            breakRowHeightScale : 0.75,
            lineWidth : 0.65,
            emptyLineWidth : 0.5,
            fonts : {
                title : 15,
                dayHeader : 12,
                weekHeader : 8,
                time : 10,
                lesson : 12,
                lessonLineHeight : 9.8
            },
            lessonPadding : {
                x : 4,
                y : 6
            }
        }
    },
}
