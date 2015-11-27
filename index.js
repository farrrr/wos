var fs = require('fs');
var xlsx = require('node-xlsx');
var _ = require('lodash');

var mappingFile = xlsx.parse(__dirname + '/data/學校老師英文對照數據庫_1026.xlsx');
var wosFile = xlsx.parse(__dirname + '/data/1125-2010-2014-5825-2.xls');

var mappingDb = _.drop(_.map(mappingFile[0].data, function(data) {
  data = _.map(data, _.trim);

  return {
    cname: data[0],
    ename: data[1],
    depart: data[2],
  };
}));

var finalData = _.map(wosFile[0].data, function(data) {
  var authors = _.map(data[1].split(';'), _.trim);
  var cauthors = [];
  var departs = [];
  _.each(authors, function(author) {
    var matched = _.where(mappingDb, { ename: author });
    var isDuplicate = matched.length > 1;

    _.each(matched, function(match) {
      cauthors.push((isDuplicate ? '$' : '') + match.cname + '(' + match.depart + ')');
      departs.push((isDuplicate ? '$' : '') + match.depart);
    });
  });

  data.push(cauthors.join(';'));
  data.push(departs.join(';'));

  return data;
});

fs.writeFileSync(__dirname + '/data/final.xlsx', xlsx.build([{ name: 'final', data: finalData }]), 'binary');
