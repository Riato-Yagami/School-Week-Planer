global.__basedir = __dirname;
global.config = require(__basedir + '/config/config');
global.fun = require(__basedir + '/loaders/load-functions');

require(__basedir + '/loaders/load-objects.js');

const main = () => {
    fun.generateSpreadsheet(); // Appelle la fonction pour générer le fichier ODS
};

const args = process.argv.slice(2);

main(args);