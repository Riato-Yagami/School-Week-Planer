const fs = require('fs');
const path = require('path');

module.exports = () => {
    return loadClassesSchedules()
};

function loadClassesSchedules() {
    const classSchedules = [];

    // Read all files in the input folder
    const files = fs.readdirSync(path.join(config.input,'classes'));

    // Loop through each file in the folder
    files.forEach(file => {
    const filePath = path.join(config.input,'classes', file);

    // Check if the file is a .js file
    if (file.endsWith('.js')) {
        const className = path.basename(file, '.js'); // Remove .js extension to get class name
        try {
        // Dynamically require each class schedule file
        const schedule = require(path.join(__basedir,filePath));
        schedule.name = className
        // Store the schedule data in the classSchedules object
        classSchedules.push(schedule);
        } catch (err) {
        console.error(`Error loading schedule for class ${className}:`, err);
        }
    }
    });

    return classSchedules;
}