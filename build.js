const uglify = require('uglify-js');
const path = require('node:path');
const fs = require('node:fs/promises');

(async () => {

    const tsconfig = require('./tsconfig.json');
    const dist = path.join(__dirname, tsconfig.compilerOptions.outDir);
    const files = (await fs.readdir(dist)).filter(file => /\.js$/i.test(file));
    const code = {};

    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const filepath = path.join(dist, file);
        code[file] = (await fs.readFile(filepath)).toString('utf-8');
    }

    const results = uglify.minify(code);

    if (results.error)
        return console.error(results.error);

    const output = path.join(__dirname, 'build');
    const outputpath = path.join(output, 'build.min.js');

    try {
        await fs.rm(output);
    } catch (ex) {
    }

    try {
        await fs.mkdir(output);
    } catch (ex) {
    }

    await fs.writeFile(outputpath, results.code);

})();
