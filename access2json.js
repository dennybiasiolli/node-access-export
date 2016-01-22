var ADODB = require('node-adodb');
var fs = require('fs');

var pathDB = process.argv[2];

var connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + pathDB + ';');
ADODB.debug = false;

/// ottiene la lista di tutte le tabelle del db, però non funziona!
//SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0

// test command
//node access2json ./Mud2015.mdb Attivita Codici CodRegCE ComConv Comuni Nazioni Province Rifiuti

for(var i = 3; i < process.argv.length; i++){
    exportTable(process.argv[i]);
}


function exportTable(_tableName){
    connection.query(
        'SELECT * FROM ' + _tableName
    ).on(
        'done',
        function (data){
            var exportName = pathDB.replace('.mdb', '') + '.' + _tableName + '.json';
            //console.log(JSON.stringify(data.records, null, '  '));
            console.log('Exporting ' + exportName);
            fs.writeFileSync(exportName, JSON.stringify(data.records, null, '  '));
        }
    ).on(
        'fail',
        function (data){
            // TODO ???? 
        }
    );
}
