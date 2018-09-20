function archive_project_on_selected_row(){
  return archive_project( SpreadsheetApp.getActiveRange().getRow() );
}

function archive_project(row_number){
  var active_sheet = SpreadsheetApp.getActiveRange().getSheet().getName();
  //make sure we are on the right sheet
  if( active_sheet !== "Projects" ){
    var error_msg = "ERROR! Invalid sheet '"+ active_sheet +"' selected. Must have 'Projects' Sheet selected.";
    Logger.log(error_msg);
    Browser.msgBox(error_msg);
    return false;
  }
  
  var now = new Date();
  var proj_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Projects");
  
  //make sure we are on a valid row that is not header or formula row(1 & 2)
  if( row_number <= 2 || row_number > proj_sheet.getLastRow() ){
    var error_msg = "ERROR! Invalid row_number '"+ row_number +"' selected. Must have a row greater than 2 selected.";
    Logger.log(error_msg);
    Browser.msgBox(error_msg);
    return false;
  }
  
  var archive_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Projects_Archive");
  var row_values = proj_sheet.getRange(row_number, 1, 1, proj_sheet.getLastColumn()).getValues();
  row_values = row_values.shift();
  //prepend archive time
  row_values.unshift(now);
  //Logger.log( JSON.stringify(row_values) );
  //append row to archive sheet
  archive_sheet.appendRow(row_values);
  
  //delete row from project sheet
  proj_sheet.deleteRow(row_number);
  
  //fix formatting
  var archive_sheet_row_number = archive_sheet.getLastRow();
  archive_sheet.getRange( archive_sheet_row_number,  1, 1, 1).setNumberFormat('yyyy"-"mm"-"dd"   "hh":"mm":"ss'); //Archive Date
  archive_sheet.getRange( archive_sheet_row_number, 12, 1, 1).setNumberFormat('yyyy"-"mm"-"dd'); //Production Start Date
  archive_sheet.getRange( archive_sheet_row_number, 14, 1, 1).setNumberFormat('yyyy"-"mm"-"dd'); //Ship Date
  archive_sheet.getRange( archive_sheet_row_number, 16, 1, 1).setNumberFormat('yyyy"-"mm"-"dd'); //Receiving Date
  archive_sheet.getRange( archive_sheet_row_number, 17, 1, 1).setNumberFormat('yyyy"-"mm"-"dd'); //Install Date
  archive_sheet.getRange( archive_sheet_row_number, 24, 1, 1).setNumberFormat('yyyy"-"mm"-"dd'); //Material arrival Cut Off Date
    archive_sheet.getRange( archive_sheet_row_number, 35, 1, 1).setNumberFormat('yyyy"-"mm"-"dd'); //Quote Submitted Date
    archive_sheet.getRange( archive_sheet_row_number, 36, 1, 1).setNumberFormat('yyyy"-"mm"-"dd'); //Quote Approved Date
    archive_sheet.getRange( archive_sheet_row_number, 38, 1, 1).setNumberFormat('yyyy"-"mm"-"dd'); //Material Missing Notify Date
  
  //Done!
  return true;
}

function get_store_info(store_id){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Store Info");
  var store_ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues();
  var store_id_count = store_ids.length;
  for( var i = 0; i < store_id_count ; i++){
    if( store_ids[ i ][0] == store_id ){
      var store_row = i + 2;
          break;
    }
  }
  
  Logger.log(store_row);
  Logger.log(store_id);
  var row_data = sheet.getRange( 0+ store_row, 1, 1, sheet.getMaxColumns()).getValues();
  var col_map = get_col_maps("Store Info");
  
  //Logger.log(row_data);
  //Logger.log(col_map);
  
  var result = {};
  for( var i = 0 ; i < row_data[0].length ; i++){
    var column_name = col_map.id_to_name[ i+1 ];
    var value = row_data[0][ i ];
    result[ column_name ] = value;
  }
  return result;
}

function log_change(change){
  var log_sheet = SpreadsheetApp.openById(log_sheet_key).getSheetByName("log");
  
  var new_row = [
    change.date,
    change.sheet_name,
    change.edit_set_a1,      
    change.row_id,
    change.col_id,
    change.column_name,
    change.user,
    change.old_value,
    change.value
  ];
  
  if( change.sheet_name === 'Projects' && typeof(change.store_id)!=='undefined' ){
    new_row.push( change.store_id );
    new_row.push( change.store_name );
    new_row.push( change.project_description );
    new_row.push( change.trello_card_id );
  }
  //log
  var last_row_num = log_sheet.appendRow(new_row).getLastRow();
  //fix date format
  
  log_sheet.getRange( last_row_num, 1, 1, 1).setNumberFormat('yyyy"-"mm"-"dd HH":"mm":"ss');
  
    //fix date format for value/old_value if this change was a column with 'date' in the name
  if( change.column_name.toLowerCase().indexOf( 'date' ) != -1 || change.column_name==='Grand Opening' ){
    log_sheet.getRange(last_row_num, 8, 1, 2).setNumberFormat('yyyy"-"mm"-"dd');
  }
}

function myOnEditOld(e){
  var range = e.range;
  
  var column_start = range.getColumn();
  var column_stop = range.getLastColumn();
  var column_count = range.getNumColumns();
  var row_start = range.getRow();
  var row_stop = range.getLastRow();
  var row_count = range.getNumRows();
  
  var sheet_name = range.getSheet().getName();
  
  var col_map = get_col_maps(sheet_name);
  var log_sheet = SpreadsheetApp.openById(log_sheet_key).getSheetByName("log");
  var test_sheet = SpreadsheetApp.openById(log_sheet_key).getSheetByName("testing");
  
  var values = range.getValues();
  var num_rows = values.length;
  var num_cols = values[0].length;
  var changes = [];
  var now = new Date();
  var sheet_col_count = range.getSheet().getMaxColumns();
  
  
  for( var i = 0; i < num_rows; i++){
    if( sheet_name === 'Projects' ){
          //gets first cell in edited row
      var edited_row = range.getSheet().getRange(row_start+i, 1, 1, range.getSheet().getMaxColumns() ).getValues()[0];
    }
    if( num_cols===sheet_col_count && values[i].join('')==='' ){
      //blank row inserted
      var change = {};
      change['date']      = now;
      change['sheet_name']  = sheet_name;
      change['edit_set_a1']   = range.getA1Notation();
      change['row_id']    = i + row_start;
      change['col_id']    = '*';
      change['column_name'] = '*';
      change['user']      = e.user.getEmail();
      change['value']     = "<<<ROW_INSERT>>";
      log_change(change);
      changes.push(change);
    }else{
      if( sheet_name === 'Projects' ){
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Projects");
        var store_id = edited_row[ col_map.name_to_id['Store ID']-1 ]
        if( store_id ){
          var store_info = get_store_info( store_id );
          
          if( store_info['Store Name'] != edited_row[ col_map.name_to_id['Store Name']-1 ] ){
            sheet.getRange( i+row_start, col_map.name_to_id['Store Name'], 1, 1).setValue( store_info['Store Name'] );
          }
          if( store_info['Address String'] != edited_row[ col_map.name_to_id['Address']-1 ] ){
            sheet.getRange( row_start+0 , col_map.name_to_id['Address'], 1, 1).setValue( store_info['Address String'] );
          }
          if( store_info['Market']!='' && store_info['Market'] != edited_row[ col_map.name_to_id['Market']-1 ] ){
            sheet.getRange( row_start+0 , col_map.name_to_id['Market'], 1, 1).setValue( store_info['Market'] );
          }
          if( store_info['Region']!='' && store_info['Region'] != edited_row[ col_map.name_to_id['Region']-1 ] ){
            sheet.getRange( row_start+0 , col_map.name_to_id['Region'], 1, 1).setValue( store_info['Region'] );
          }
          
          var project_description = edited_row[ col_map.name_to_id['Description']-1 ];
          if( project_description ) {
            var trello_card_name_current = edited_row[ col_map.name_to_id['Trello Card Name']-1 ];
            var trello_card_name_expected = store_id +" "+ store_info['Store Name'] +" - "+project_description;
            if( trello_card_name_current != trello_card_name_expected ){
              sheet.getRange( i+row_start, col_map.name_to_id['Trello Card Name'], 1, 1).setValue(trello_card_name_expected);
            }
          }
                  var ship_date = edited_row[ col_map.name_to_id['Ship Date']-1 ];
                  if( ship_date ){
                    var macod_current = sheet.getRange(row_start + i, 0+ col_map.name_to_id['Material Arrival Cut Off Date'] ).getFormula();
                    var macod_expected = "=if(isblank(M"+ (row_start+i) +"),,WORKDAY( M"+ (row_start+i) +",-10))";
                    if( macod_current != macod_expected ){
                      sheet.getRange( row_start+i , col_map.name_to_id['Material Arrival Cut Off Date'], 1, 1).setFormula(macod_expected).setNumberFormat('yyyy"-"mm"-"dd');
                    }
                    //Material Missing Notify Date
                    var mmnd_current = sheet.getRange(row_start + i, 0+ col_map.name_to_id['Material Missing Notify Date'] ).getFormula();
                    var mmnd_expected = "=if(isblank(M"+ (row_start+i) +"),,WORKDAY( M"+ (row_start+i) +",-15))";
                    if( mmnd_current != mmnd_expected ){
                      sheet.getRange( row_start+i , col_map.name_to_id['Material Missing Notify Date'], 1, 1).setFormula(mmnd_expected).setNumberFormat('yyyy"-"mm"-"dd');
                    }
                    
                    var dl_current = sheet.getRange(row_start + i, 0+ col_map.name_to_id['Days Left'] ).getFormula();
                    var dl_expected = "=if(isblank(M"+ (row_start+i) +"),, if(eq(B"+ (row_start+i) +",\"Complete\"),, if(eq(B"+ (row_start+i) +",\"Shipped\"),, if(eq(B"+ (row_start+i) +", \"Shipped - Partial\"),, M"+ (row_start+i) +"-today() ))))";
                    if( dl_current != dl_expected ){
                      sheet.getRange( row_start+i , col_map.name_to_id['Days Left'], 1, 1).setFormula(dl_expected);
                    }
                    var dbsa_current = sheet.getRange(row_start + i, 0+ col_map.name_to_id['Number of Days Between Submitted & Approved'] ).getFormula();
                    var dbsa_expected = "=if(isblank(AJ"+ (row_start+i) +"),, if(isblank(AK"+ (row_start+i) +"),, round(AK"+ (row_start+i) +" - AJ"+ (row_start+i) +")))";
                    if( dbsa_current != dbsa_expected ){
                      sheet.getRange( row_start+i , col_map.name_to_id['Number of Days Between Submitted & Approved'], 1, 1).setFormula(dbsa_expected);
                    }
                  }
        }
      }
      
      for( var j = 0; j < num_cols; j++){
        var change = {};
        change['date']      = now;
        change['sheet_name']  = sheet_name;
        change['edit_set_a1']   = range.getA1Notation();
        change['row_id']    = i + row_start;
        change['col_id']    = j + column_start;
        change['column_name'] = col_map.id_to_name[ j+column_start ];
        change['user']      = e.user.getEmail();
        change['value']     = values[i][j];
        
        //oldValue is only set if the edit was a single cell
        if( num_rows == 1 && num_cols == 1 ){
          change['old_value'] = e.oldValue;
        }
        
        if( sheet_name === 'Projects' ){
          change['store_id'] = store_id;
          change['store_name'] = store_info['Store Name'];
          change['project_description'] = project_description;
          change['trello_card_id'] = edited_row[ col_map.name_to_id['Trello Card ID']-1 ];
          if( change['column_name'] === 'Status' && change['value'] === 'In Production' ){
                      var now = new Date();
            sheet.getRange(row_start, col_map.name_to_id['Production Start Date'], 1, 1).getCell(1,1).setNumberFormat('yyyy"-"mm"-"dd').setValue(now);
          }
        }
        
        log_change(change);
        changes.push(change);
      }
    }
  }
  
}

function get_col_maps(sheet_name, row_no){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  
  if( ! sheet ){
    return false;
  }
  
  var header_row_cols = sheet.getRange( sheet.getFrozenRows()||1 , 1, 1, sheet.getMaxColumns()).getValues();
  
  //Logger.log("header_row_cols");
  //Logger.log(header_row_cols);
  
  var cm_id_to_name = [];
  var cm_name_to_id = {};
  
  for( var i = 0, length = header_row_cols[0].length; i < length; i++){
    cm_id_to_name[i + 1] = header_row_cols[0][i];
    
    //only set the first occurrence of each unique string
    if(typeof(cm_name_to_id[header_row_cols[0][i]]) === 'undefined' ){
      cm_name_to_id[ header_row_cols[0][i] ] = i+1;
    }
  }
  
  return {
    'id_to_name':cm_id_to_name,
    'name_to_id':cm_name_to_id
  };
}

function insert_time() {
  var now = new Date();
  SpreadsheetApp.getActiveSheet().getActiveCell().setValue(now);
}
