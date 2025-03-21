//check for Japanese text in files folder and save them to an Excel file
//Usage: node tool.js <folder_path>
//Example: node tool.js /Users/username/Documents/Project
//required libraries: glob, cheerio, japanese, xlsx, node > v15.0.0
//npm install glob cheerio japanese xlsx

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

    // Include .vue, .js, and .ts files in the glob pattern
    const filesPath = path.join(folderPath, '**/*.{vue,js,ts}');
    let foundTexts = [];
    
    function checkFile(filePath) {
        const content = fs.readFileSync(filePath, 'utf-8');
        const $ = cheerio.load(content, { xmlMode: true });
        let fileTexts = new Set();

        $('*').each((index, element) => {
            const text = $(element).text().trim();
            const japaneseMatches = text.match(/[\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Han}]+/gu);
            if (japaneseMatches) {
                japaneseMatches.forEach(match => {
                    fileTexts.add(match);
                });
            }
        });

        if (fileTexts.size > 0) {
            foundTexts.push({ File: filePath, JapaneseTexts: Array.from(fileTexts) });
        }
    }

    const files = await glob(filesPath);
    files.forEach(file => checkFile(file));

    if (foundTexts.length > 0) {
        saveToExcel(foundTexts);
    } else {
        console.log('No Japanese text found.');
    }
}

function saveToExcel(data) {
    let formattedData = [];
    data.forEach(item => {
        let row = { File: item.File };
        item.JapaneseTexts.forEach((text, index) => {
            row[`text${index + 1}`] = text;
        });
        formattedData.push(row);
    });

    const ws = xlsx.utils.json_to_sheet(formattedData);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Japanese Texts');

    let fileName = 'Japanese_Texts.xlsx';
    let counter = 1;
    while (fs.existsSync(fileName)) {
        fileName = `Japanese_Texts_${counter}.xlsx`;
        counter++;
    }

    xlsx.writeFile(wb, fileName);
    console.log(`Japanese texts saved to ${fileName}`);
}

const folderPath = process.argv[2];

if (!folderPath) {
    console.error('Please provide a folder path as an argument.');
    process.exit(1);
}

installLibraries();
runTool(folderPath);