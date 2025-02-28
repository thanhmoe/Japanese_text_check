const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const requiredLibraries = ['glob', 'cheerio', 'japanese', 'xlsx'];

function installLibraries() {
    requiredLibraries.forEach(lib => {
        try {
            require.resolve(lib);
        } catch (e) {
            console.log(`Installing ${lib}...`);
            execSync(`npm install ${lib}`, { stdio: 'inherit' });
        }
    });
}

async function runTool(folderPath) {
    const { glob } = require('glob');
    const cheerio = require('cheerio');
    const japanese = require('japanese');
    const xlsx = require('xlsx');

    const vueFilesPath = path.join(folderPath, '**/*.vue');
    let foundTexts = [];
    
    function checkVueFile(filePath) {
        const content = fs.readFileSync(filePath, 'utf-8');
        const $ = cheerio.load(content, { xmlMode: true });

        $('*').each((index, element) => {
            const text = $(element).text().trim();
            const japaneseMatches = text.match(/[\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Han}]+/gu);
            if (japaneseMatches) {
                japaneseMatches.forEach(match => {
                    foundTexts.push({ File: filePath, JapaneseText: match });
                });
            }
        });
    }

    const files = await glob(vueFilesPath);
    files.forEach(file => checkVueFile(file));

    if (foundTexts.length > 0) {
        saveToExcel(foundTexts);
    } else {
        console.log('No Japanese text found.');
    }
}

function saveToExcel(data) {
    const ws = xlsx.utils.json_to_sheet(data);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Japanese Texts');
    xlsx.writeFile(wb, 'Japanese_Texts.xlsx');
    console.log('Japanese texts saved to Japanese_Texts.xlsx');
}

const folderPath = process.argv[2];

if (!folderPath) {
    console.error('Please provide a folder path as an argument.');
    process.exit(1);
}

installLibraries();
runTool(folderPath);