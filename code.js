function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menu_items = [
    { name: 'Sync with Trello', functionName: 'sync_with_trello' }, 
    { name: 'Archive Project on Selected Row', functionName: 'archive_project_on_selected_row' },
    { name: 'Insert Time', functionName: 'insert_time' }  
  ];
  spreadsheet.addMenu('Scripts', menu_items);
}

function onEdit(e) {
  var r = e.range;
  Logger.log(JSON.stringify(e));
  if(r.getSheet().getName() === 'Projects') {
    if(r.getNumColumns() === r.getSheet().getMaxColumns() && !r.isBlank()) {
      // Copied entire row(s)
      for(var i = r.getRow() -1; i < r.getLastRow(); i++) {
        var o = {};
        for(var k in projects) {
          if(projects[k].a === 'blank') projects[k].r.offset(i, 0, 1, 1).setValue('');
          if(projects[k].a === 'formula') projects[k].r.offset(i, 0, 1, 1).setFormula(projects[k].f);
          if(projects[k].a === 'code' && projects[k].n === 'store_po_no') projects[k].r.offset(i, 0, 1, 1).setValue(projects[4].r.offset(i, 0).getValue() + '-' + projects[3].r.offset(i, 0).getValue());
          if(projects[k].a === 'code' && projects[k].n === 'row_timestamp') projects[k].r.offset(i, 0, 1, 1).setValue((new Date()).getTime());
          o[projects[k].n] = projects[k].r.offset(i, 0, 1, 1).getValue();
        }
        queue_sync(o);
      } 
    } else if(r.getNumColumns() === r.getSheet().getMaxColumns() && r.isBlank()) {
      // Do nothing
    } else if(r.getNumRows() === 1 && r.getNumColumns() === 1) {
      // Only one cell changed
      var o = {};
      if(get('n', 'row_timestamp')[0].r.getColumn() === r.getColumn()) {
          r.setValue(typeof e.oldValue === 'undefined' ? '' : e.oldValue);
          SpreadsheetApp.getUi().alert('ROW_TIMESTAMP Don\'t manually update this columns. Your last action was undone to avoid breaking the sync with Trello.')
        } else if(!get('n', 'row_timestamp')[0].r.offset(r.getRow() -1, 0, 1, 1).getValue()) {
          get('n', 'row_timestamp')[0].r.offset(r.getRow() -1, 0, 1, 1).setValue((new Date()).getTime());
        }
      for(var k in projects) o[projects[k].n] = projects[k].r.offset(r.getRow() -1, 0, 1, 1).getValue();
      queue_sync(o);
    } else {
      // Partial row(s) copied
      if(get('n', 'row_timestamp')[0].r.getColumn() >= r.getColumn() && get('n', 'row_timestamp')[0].r.getColumn() <= r.getLastColumn()) {
          SpreadsheetApp.getUi().alert('ROW_TIMESTAMP Don\'t manually update this columns. Please undo your last action to avoid breaking the sync with Trello.')
        } else {
          for(var i = r.getRow() -1; i < r.getLastRow(); i++) {
            var o = {};
            if(!get('n', 'row_timestamp')[0].r.offset(i, 0, 1, 1).getValue()) get('n', 'row_timestamp')[0].r.offset(i, 0, 1, 1).setValue((new Date()).getTime());
            for(var k in projects) o[projects[k].n] = projects[k].r.offset(i, 0, 1, 1).getValue();
            queue_sync(o);
          }
        }
    }
  }
}

function doGet(e) {
  var req = e.parameters;
  var res = false;
  
  if(typeof req.operation !== 'undefined' && (typeof req.row_timestamp !== 'undefined' || typeof req.trello_card_id !== 'undefined')) {
    if(req.operation.toString().indexOf('sanitation') === -1) {
      var sync_headers = sync.getRange('1:1').getValues();
      var ref_sync_col = typeof req.row_timestamp !== 'undefined' ? 'ROW_TIMESTAMP' : 'TRELLO_CARD_ID';
      ref_sync_col = sync_headers[0].indexOf(ref_sync_col) +1;
      var update_sync_cols = ['TRELLO_CARD_ID', 'STATUS', 'LAST_COMMIT'];
      update_sync_cols = update_sync_cols.map(function(c) { return sync_headers[0].indexOf(c) +1 });
      var arr = sync.getRange(1, ref_sync_col, sync.getLastRow(), 1).getValues();
      arr = transpose_col_values(arr);
      arr = find_row_numbers_in_col(arr, (typeof req.row_timestamp !== 'undefined' ? req.row_timestamp : req.trello_card_id));
      if(arr.length === 1) {
        if(typeof req.trello_card_id !== 'undefined' && req.operation == 'new_card') sync.getRange(arr[0], update_sync_cols[0]).setValue(req.trello_card_id);
        //sync.getRange(arr[0], 3).clearContent();
        sync.getRange(arr[0], update_sync_cols[1]).setValue(req.sync_value);
        sync.getRange(arr[0], update_sync_cols[2]).setValue((new Date()).getTime());
        if(typeof req.row_timestamp === 'undefined') {
          ref_sync_col = sync_headers[0].indexOf('ROW_TIMESTAMP') +1;
          req.row_timestamp = sync.getRange(arr[0], ref_sync_col, 1, 1).getValue();
        }
        res = true;
      } else {
        // sync is messed up || older card || card created in trello
      }
    }
    
    if(req.operation.toString() === 'trello_move' || req.operation.toString() === 'trello_update') {
      var ref_cols = [projects.filter(function(o) { return o.n === 'row_timestamp' })[0], projects.filter(function(o) { return o.n === 'trello_card_name' })[0]];
      var ref_row = null;
      for(var i in ref_cols) {
        ref_row = find_row_numbers_in_col(transpose_col_values(ref_cols[i].r.getValues()), req[ref_cols[i].n]);
        if(ref_row.length > 0) break;
      }
      if(ref_row.length === 1) {
        var update_cols = ['status', 'shipping_date', 'trello_card_name'];
        for(var i in update_cols) {
          if(typeof req[update_cols[i]] !== 'undefined') {
            var update_range = projects.filter(function(o) { return o.n === update_cols[i] })[0].r;
            var update_value = update_cols[i] === 'shipping_date' ? Utilities.formatDate(new Date(req[update_cols[i]]), 'GMT', 'yyyy-MM-dd') : req[update_cols[i]];
            update_range.offset((ref_row[0] -1), 0, 1, 1).setValue(update_value);
          }
        }
      } else {
        if(res) sync.getRange(arr[0], update_sync_cols[1]).setValue('SYNC FAILED');
        res = false;
      }
    }
    
    if(req.operation.toString().indexOf('sanitation') > -1) {
      var new_row = [];
      new_row.push(typeof req.row_timestamp === 'undefined' ? '' : req.row_timestamp[0]);
      new_row.push(typeof req.trello_card_id === 'undefined' ? '' : req.trello_card_id[0]);
      new_row.push(typeof req.trello_card_name === 'undefined' ? '' : req.trello_card_name[0]);
      new_row.push(typeof req.status === 'undefined' ? '' : req.status[0]);
      new_row.push(typeof req.shipping_date === 'undefined' ? '' : req.shipping_date[0]);
      new_row.push(typeof req.issue === 'undefined' ? '' : req.issue[0]);
      new_row.push(typeof req.id === 'undefined' ? '' : req.id[0]);
      new_row.push(typeof req.list_name === 'undefined' ? '' : req.list_name[0]);
      new_row.push(typeof req.due === 'undefined' ? '' : req.due[0]);
      SpreadsheetApp.getActive().getSheetByName('sanitation').appendRow(new_row);
      res = true;
    }
  }
  return ContentService.createTextOutput('{ "ok": ' + res + ' }');
}
