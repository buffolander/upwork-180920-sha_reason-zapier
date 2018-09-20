function sanitation() {
  var s = ss.getSheetByName('Projects');
  /* 
    // Find exact duplicate rows
    var r = s.getRange(1, 1, s.getLastRow(), s.getLastColumn()).getValues();
    var col_array = [1, 2, 3, 4, 5, 7, 8, 9, 10, 12, 13, 16, 17];
    var d_rows = compare_function(r, col_array, false, 'number');
    Logger.log(d_rows);
    // Delete duplicate rows
    for(var i = d_rows.length -1; i >= 0; i--) s.deleteRow(d_rows[i]);
  DONE Duplicate Rows
  */
  /*
    // Set row_timestamps
    var r = s.getRange(1, 1, s.getLastRow(), s.getLastColumn()).getValues();
    var col_array = [18];
    var d_rows = compare_function(r, col_array, true, 'number');
    for(i in d_rows) {
      getProjectCol('row_timestamp').r.offset(d_rows[i] -1, 0, 1, 1).setValue((new Date()).getTime());
      Utilities.sleep(5);
    }
  DONE Set row_timestamps
  */
  /*
    // Check trello_card_id orphans
    var r = s.getRange(1, 1, s.getLastRow(), s.getLastColumn()).getValues();
    var col_array = [17];
    var e_rows = r.map(function(v, i ) { return (v[col_array[0] -1].toString().trim() === '' && v[0] !== '') ? v : -1 }).filter(function(v) { return v !== -1 });
    Logger.log(e_rows.length);
    for(i in e_rows) {
      var trello_query = { "row_timestamp": "", "trello_card_id": "", "trello_card_name": "", "status": "", "shipping_date": "", "issue": "MISSING" };
      trello_query.row_timestamp = e_rows[i][17];
      trello_query.trello_card_name = e_rows[i][15];
      trello_query.status = e_rows[i][1];
      trello_query.shipping_date = e_rows[i][11];
      request_card_details(trello_query);
    }
  DONE Check trello_card_id orphans
  */
  /*
    // Check duplicate trello_card_ids
    var r = s.getRange(1, 1, s.getLastRow(), s.getLastColumn()).getValues();
    var col_array = [17];
    var d_rows = compare_function(r, col_array, true, 'values');
    Logger.log(d_rows.length);
    for(i in d_rows) {
      var trello_query = { "row_timestamp": "", "trello_card_id": "", "trello_card_name": "", "status": "", "shipping_date": "", "issue": "DUPLICATE" };
      trello_query.row_timestamp = d_rows[i][17];
      trello_query.trello_card_id = d_rows[i][16];
      trello_query.trello_card_name = d_rows[i][15];
      trello_query.status = d_rows[i][1];
      trello_query.shipping_date = d_rows[i][11];
      request_card_details(trello_query);
      Utilities.sleep(1000);
    }
  DONE Check deuplicate trello_card_id
  */
  /*
    // Wrap-up Sanitation
    var r = s.getRange(1, 1, s.getLastRow(), s.getLastColumn()).getValues();
    var r_sanitation = SpreadsheetApp.getActive().getSheetByName('sanitation');
    r_sanitation = r_sanitation.getRange(4, 1, r_sanitation.getLastRow(), r_sanitation.getLastColumn()).getValues();
    for(var i in r_sanitation) {
      var row = r.map(function(v, index) { return v[17] === r_sanitation[i][0] ? index : -1 }).filter(function(v) { return v !== -1 });
      if(row.length > 0) {
        row = row[0];
        // IF trello_card_id !== CARDNOTFOUND 
        // update sheet 'Projects' trello_card_id returned form Trello
        if(r_sanitation[i][6] !== 'CARDNOTFOUND') s.getRange(row +1, 17, 1, 1).setValue(r_sanitation[i][6]);
        else s.getRange(row +1, 17, 1, 1).setValue('');
        // Queue Create/Update Card
        var o = {};
        for(var k in projects) o[projects[k].n] = projects[k].r.offset(row, 0, 1, 1).getValue();
        queue_sync(o);
      }
    }
    var check_sync = transpose_col_values(sync.getRange('A:A').getValues())[0];
    r = r.map(function(v) { return check_sync.indexOf(v[17]) === -1 && v[16].trim() !== '' ? v : -1 }).filter(function(v) { return v !== -1 });
    for(var i in r) sync.appendRow([r[i][17], r[i][16]]);
  DONE Wrap-up
  */
}

function compare_function(haystack, col_array, include_current, row) {
  var d_rows_set = [];
  for(var i = 0; i < haystack.length; i++) {
    //var d_rows = d_rows_set.indexOf(i +1) > -1 ? [] : compare_rows(haystack[i], haystack, i, col_array);
    var d_rows = (include_current || d_rows_set.indexOf(i +1) === -1) ? compare_rows(haystack[i], haystack, i, col_array, row) : [];
    if(d_rows.length > 0) d_rows_set = d_rows_set.concat(d_rows);
  }
  var res = [];
  for(var i in d_rows_set) if(res.indexOf(d_rows_set[i]) === -1) res.push(d_rows_set[i]);
  return res.sort(function(a, b) { return a - b });
}

function compare_rows(needle, haystack, index, col_array, row) { // col array use col numbers, not indexes
  return haystack
  .map(function(v, i) {
    var test = i !== index ? true : false;
    var blank = true;
    if(test) 
      for(j in col_array) {
        test = v[col_array[j] -1].toString() !== needle[col_array[j] -1].toString() ? false : test; 
        blank = v[col_array[j] -1].toString().trim() !== '' ? false : blank;
      }
    return (test && !blank) ? (row === 'number' ? i +1 : v) : -1;
  })
  .filter(function(v) {
    return v !== -1;
  });
}

function group_by(collection, property) {
  var i = 0, values = [], result = [];
  for (i; i < collection.length; i++) {
    if(values.indexOf(collection[i][property]) === -1) {
      values.push(collection[i][property]);
      result.push(collection.filter(function(v) { return v[property] === collection[i][property] }));
    }
  }
  return result;
}

function request_card_details(args) {
  var url = 'https://hooks.zapier.com/hooks/catch/3731004/qnvszc/';
  var options = { 
    "method": "post", 
    "payload": args
  };
  if(url) res = UrlFetchApp.fetch(url, options);
}
