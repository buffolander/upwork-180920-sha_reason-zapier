function create_named_ranges() {
  var r = SpreadsheetApp.getActive().getRange('Projects!1:1');
  for(var i = 1; i <= r.getNumColumns(); i++) {
    var n = 'Projects_' + r.getCell(1, i).getValue();
    var nr = r.getCell(1, i).getA1Notation().match(/[A-Z]+/g)[0];
    nr = 'Projects!' + nr + ':' + nr;
    SpreadsheetApp.getActive().setNamedRange(n, SpreadsheetApp.getActive().getRange(nr));
  }
}

function get(k, q) { // filter project field(s) details
  var allowed_k = ['n', 'a'];
  if(allowed_k.indexOf(k) !== -1) return projects.filter(function(o) { return o[k] == q });
  else return false;
}

function queue_sync(row) {
  if(row.project && row.status && row.product && row.store_id) {
    var arr = sync.getRange('A:A').getValues();
    arr = Object.keys(arr[0]).map(function(c) { return arr.map(function (r) { return r[c]; }) });
    arr = arr[0].map(function(v, i) { return v === row.row_timestamp ? i +1 : -1 }).filter(function(v) { return v > -1 });
    if(arr.length === 0) sync.appendRow([row.row_timestamp, row.trello_card_id, 'CREATE', 'PENDING', JSON.stringify(row)]);
    if(arr.length > 0) {
      var trello_card_id = sync.getRange(arr[arr.length -1], 2).getValue();
      trello_card_id = trello_card_id !== 'undefined' ? trello_card_id : '';
      row.trello_card_id = trello_card_id;
      sync.getRange(arr[arr.length -1], 1, 1, 5).setValues([[row.row_timestamp, trello_card_id, trello_card_id !== '' ? 'UPDATE': 'CREATE', 'PENDING', JSON.stringify(row)]]);
    }
    if(arr.length > 1) for(var i = 0; i < arr.length; i++) sync.deleteRow(arr[i]);
  }
}

function sync_with_trello() {
  var arr = sync.getRange('A:E').getValues();
  arr = arr.filter(function(a) { return a[3] === 'PENDING' });
  //arr = Object.keys(arr[0]).map(function(c) { return arr.map(function (r) { return r[c]; }) });
  Logger.log(arr);
  for(var i in arr) {
    var url = res = false;
    if(arr[i][2] == 'CREATE') url = 'https://hooks.zapier.com/hooks/catch/3731004/qsi0fe/';
    if(arr[i][2] == 'UPDATE') url = 'https://hooks.zapier.com/hooks/catch/3731004/q885iq/';
    var options = { "method": "post", "payload": arr[i][4] };
    if(url) res = UrlFetchApp.fetch(url, options);
  }
}

function transpose_col_values(arr) {
  return Object.keys(arr[0]).map(function(c) { return arr.map(function (r) { return r[c]; }) });
}

function find_row_numbers_in_col(arr, key) {
  return arr[0].map(function(v, i) { return v.toString() === key.toString() ? i +1 : -1 }).filter(function(v) { return v > -1 });
}
