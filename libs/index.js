var fs = require('fs');
var xlsx = require('xlsx');
var cvcsv = require('csv');

exports = module.exports = XLS_json;

// exports.XLS_json = XLS_json;

function XLS_json (config, callback) {
  if(!config.input) {
    callback(new Error("node-xls-json: You did not provide an input file."), null);
  }

  var cv = new CV(config, callback);
  
}


function CV(config, callback) { 
  var wb = this.load_xls(config.input);
  var ws = this.ws(wb, config.sheet);
  var csv = this.csv(ws);
  this.cvjson(csv, config.output, callback, config.rowsToSkip || 0, config.columns || null, config.filter || null, config.sort || null, config.columnMapper || null);
}

CV.prototype.load_xls = function(input) {
  return xlsx.readFile(input);
}

CV.prototype.ws = function(wb, target_sheet) {
  ws = wb.Sheets[target_sheet ? target_sheet : wb.SheetNames[0]];
  return ws;
}

CV.prototype.csv = function(ws) {
  return csv_file = xlsx.utils.make_csv(ws);
}


CV.prototype.cvjson = function(csv, output, callback, rowsToSkip, columns, filter , sort , columnMapper) {
  var records = [];
  var headers = [];

  cvcsv()
    .from.string(csv)
    .transform( function(row){
      row.unshift(row.pop());
      return row;
    })
    .on('record', function(row, index){
      if(index === rowsToSkip) {
        headers = row.map(function(name,index){//get headers
          return {name,index};
        });
        if(columns instanceof Array){//select headers
          headers = headers.filter(function(header){
            return columns.includes(header["name"]);
          });
        }
      }else if (index > rowsToSkip) {
        var obj = {};
        headers.forEach(function(header) {
          var key = header["name"].trim();
          var val = row[header["index"]].trim();
          obj[key] = val;
        });
        records.push(obj);
      }
    })
    .on('end', function(count){
      //extra works
      if(typeof columnMapper=="object" && columnMapper!==null){
        Object.keys(columnMapper).forEach(function(key){
          if(typeof columnMapper[key]!="function") delete columnMapper[key];//handle error arguments
        });
        records = records.map(function(record){
          for(var key in columnMapper){
            const mapper = columnMapper[key];
            var arg;
            if(record.hasOwnProperty(key)){//current column -> new value
              arg = record[key];
            } else {//current record + extra column
              arg = record;
            }
            record[key] = mapper(arg);
          }
          return record;
        });
      }
      //filter & sort after column map may cost more compute.
      //however,sometimes we need some simple mapper like Number(str) before sort
      if(typeof filter=="function") records = records.filter(filter);
      if(typeof sort=="function") records = records.sort(sort);
      // when writing to a file, use the 'close' event
      // the 'end' event may fire before the file has been written
      if(output !== null) {
      	var stream = fs.createWriteStream(output, { flags : 'w' });
      	stream.write(JSON.stringify(records));
      }
      if(typeof callback=="function") callback(null, records);
    })
    .on('error', function(error){
      if(typeof callback=="function") callback(error, null);
    });
}
