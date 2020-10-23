var electronInstaller = require('electron-winstaller');

var resultPromise = electronInstaller.createWindowsInstaller({
    appDirectory: './purgatorio-win32-x64',
    outputDirectory: './windows_installer',
    exe: './purgatorio.exe',
    title: 'Purgatorio',
    name: 'Purgatorio',
    description: 'Software para remoção de senhas de planilhas',
    authors: 'JC'
});

resultPromise.then(
    () => console.log('O instalador foi gerado com sucesso.'),
    (e) => console.log('Ocorreu um erro: ' + e.message)
);
