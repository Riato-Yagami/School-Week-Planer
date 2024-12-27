const fs = require("fs");
const path = require("path");

const dir = path.join(__basedir, "objects");

// Function to recursively load modules
function getModules(dir) {
    fs.readdirSync(dir).forEach(file => {
        const filePath = path.join(dir, file);
        const stats = fs.statSync(filePath);

        if (stats.isDirectory()) {
            // If it's a directory, recursively get modules
            getModules(filePath);
        } else if (stats.isFile() && path.extname(file) === ".js") {
            // Import the module and assign it to `global` with its filename (minus .js) as the key
            const moduleName = path.basename(file, ".js");
            global[moduleName] = require(filePath);
        }
    });
}

// Call getModules on the initial directory
getModules(dir);

// Export modules object if needed for other purposes
module.exports = global;
