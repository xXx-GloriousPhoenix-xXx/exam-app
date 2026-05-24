const inquirer = require('inquirer');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

function getVersions() {
    const scriptsDir = path.join(__dirname, 'script');
    
    if (!fs.existsSync(scriptsDir)) {
        console.error('❌ Папка /script не найдена!');
        process.exit(1);
    }
    
    const files = fs.readdirSync(scriptsDir);
    
    const versionFiles = files
        .filter(file => /^[eu]v\d+\.js$/.test(file))
        .map(file => {
            const versionNum = parseInt(file.match(/v(\d+)\.js/)[1]);
            
            const prefixType = file.startsWith('u') ? 'Uneven' : 'Even';
            
            return { 
                name: `📦 Version ${prefixType} ${versionNum}`, 
                value: file, 
                version: versionNum 
            };
        })
        .sort((a, b) => a.version - b.version);
    
    if (versionFiles.length === 0) {
        console.error('❌ Не найдено файлов версий (v1.js, v2.js и т.д.) в папке /script');
        process.exit(1);
    }
    
    const choices = [
        { name: `🔄 All (запустить все версии)`, value: 'all' },
        new inquirer.Separator('─'.repeat(30)),
        ...versionFiles.map(v => ({ name: v.name, value: v.value }))
    ];
    
    return { choices, versionFiles };
}

function runVersion(versionFile) {
    return new Promise((resolve, reject) => {
        const scriptPath = path.join(__dirname, 'script', versionFile);
        
        console.log(`🚀 Запуск ${versionFile}...`);

        const env = { ...process.env, NODE_NO_WARNINGS: '1' };
        
        const child = spawn('node', [scriptPath], {
            stdio: 'inherit',
            env: env
        });
        
        child.on('error', (error) => {
            console.error(`❌ Ошибка: ${error.message}`);
            reject(error);
        });
        
        child.on('close', (code) => {
            if (code !== 0) {
                console.error(`❌ Процесс завершился с кодом: ${code}`);
                reject(new Error(`Process exited with code ${code}`));
            } else {
                console.log(`✅ ${versionFile} завершен успешно!\n`);
                resolve();
            }
        });
    });
}

async function runAllVersions(versionFiles) {
    console.log('\n🔄 Начинаю запуск всех версий...\n');
    
    for (let i = 0; i < versionFiles.length; i++) {
        const version = versionFiles[i];
        console.log(`📌 [${i + 1}/${versionFiles.length}] ${version.name}`);
        
        try {
            await runVersion(version.value);
        } catch (error) {
            console.error(`\n❌ Ошибка при выполнении ${version.value}`);
            process.exit(1);
        }
    }
    
    console.log('='.repeat(40));
    console.log('✅ Все версии успешно выполнены!');
    console.log('='.repeat(40));
}

async function main() {
    const { choices, versionFiles } = getVersions();
    
    console.log('\n🎯 Добро пожаловать в генератор экзаменов!\n');
    
    const answer = await inquirer.prompt([
        {
            type: 'list',
            name: 'selectedVersion',
            message: 'Выберите версию для запуска:',
            choices: choices,
            pageSize: 10,
            loop: false
        }
    ]);
    
    if (answer.selectedVersion === 'all') {
        await runAllVersions(versionFiles);
    } else {
        await runVersion(answer.selectedVersion);
    }
}

process.on('SIGINT', () => {
    console.log('\n\n👋 До свидания!');
    process.exit(0);
});

main().catch(error => {
    console.error('❌ Критическая ошибка:', error);
    process.exit(1);
});