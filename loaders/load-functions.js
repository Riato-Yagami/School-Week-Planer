const fs = require("fs");
const path = require("path");

const functions_dir = path.join(__basedir, "functions");

const modules = {};

function getModules(dir) {
    fs.readdirSync(dir).forEach(file => {
        const file_path = path.join(dir, file);
        const stats = fs.statSync(file_path);
        if (stats.isDirectory()) {
            getModules(file_path);
        } else if (stats.isFile() && path.extname(file) === ".js") {
            const module = require(file_path);
            const module_mame = path.basename(file, ".js");
            modules[module_mame] = module;
        }
    });
}

getModules(functions_dir);

module.exports = modules;