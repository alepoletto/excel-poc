const Excel = require('exceljs');
const axios = require('axios');
const data = require('./data.json');
const tempfile = require('tempfile');
const fs = require('fs');
const request = require('request');

var download = function(uri, filename, callback){
  request.head(uri, function(err, res, body){
    console.log('content-type:', res.headers['content-type']);
    console.log('content-length:', res.headers['content-length']);

    request(uri).pipe(fs.createWriteStream(filename)).on('close', callback);
  });
};

let workbook = new Excel.Workbook();
let worksheet = workbook.addWorksheet('My Sheet');
worksheet.properties.defaultRowHeight = 45;

worksheet.columns = [
    { header: 'Id', key: 'id', width: 10  },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'Image', key: 'image', width: 30 }
];

const generateRows = () => {
  let promises = [];
  let count = 2;
    data.forEach((record) => {
      promises.push(new Promise((resolve, reject) => {
        let localpath = tempfile('.jpg');
        download('https://unsplash.it/200/300/?random', localpath, function(){
          var imageId1 = workbook.addImage({
                buffer: fs.readFileSync(localpath),
                extension: 'jpeg',
            });
          worksheet.addRow({id: record.id, name: record.name});
          var row = worksheet.lastRow;
          row.height = 200;
          worksheet.addImage(imageId1, `C${count}:C${count}`);
          count++;
          resolve();
        });
    }));
  });
  return Promise.all(promises);
}

generateRows().then(()=> {
  workbook.xlsx.writeFile('test.xlsx')
    .then(function() {
        // done
    });

});
