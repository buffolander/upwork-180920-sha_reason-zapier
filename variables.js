var log_sheet_key = '1BsqCYni-nv4Po7YtWv6Wua4DXscp_Oy_z1RqV21FBAs';

var ss = SpreadsheetApp.getActive();

var sync = ss.getSheetByName('projects_sync');

function getProjectCol(needle) {
  return projects.filter(function(v) { return v.n === needle })[0];
}

var projects = [
  { 
    n : "project",
    r: ss.getRangeByName('Projects_PROJECT'), 
    a: "keep"
  }, {
    n: "status",
    r: ss.getRangeByName('Projects_STATUS'), 
    a: "keep"
  }, {
    n: "product",
    r: ss.getRangeByName('Projects_PRODUCT'), 
    a: "keep"
  }, {
    n: "store_id",
    r: ss.getRangeByName('Projects_STORE_ID'), 
    a: "keep"
  }, {
    n: "store_name",
    r: ss.getRangeByName('Projects_STORE_NAME'), 
    a: "formula",
    f: "=IF(ISBLANK(INDEX(Projects_STORE_ID,ROW(),0)),,VLOOKUP(INDEX(Projects_STORE_ID,ROW(),0),'Store Info'!$A:$O,2,FALSE))"
  }, {
    n: "store_po_no",
    r: ss.getRangeByName('Projects_STORE_PO_NO'), 
    a: "code"
  }, {
    n: "description",
    r: ss.getRangeByName('Projects_DESCRIPTION'),
    a: "keep"
  }, {
    n: "fe_no",
    r: ss.getRangeByName('Projects_FE_PO'),
    a: "blank"
  }, {
    n: "leeman_sales_ord_no",
    r: ss.getRangeByName('Projects_LEEMAN_SALES_ORD_NO'),
    a: "blank"
  }, {
    n: "fi_so_id",
    r: ss.getRangeByName('Projects_FI_SO_ID'),
    a: "blank"
  }, {
    n: "days_left",
    r: ss.getRangeByName('Projects_DAYS_LEFT'),
    a: "formula",
    f: "=IF(OR(ISBLANK(INDEX(Projects_SHIPPING_DATE,ROW(),0)),EQ(INDEX(Projects_STATUS,ROW(),0),\"Complete\"),EQ(INDEX(Projects_STATUS,ROW(),0),\"Shipped\"),EQ(INDEX(Projects_STATUS,ROW(),0),\"Shipped - Partial\")),,INDEX(Projects_SHIPPING_DATE,ROW(),0)-TODAY())"
  }, {
    n: "shipping_date",
    r: ss.getRangeByName('Projects_SHIPPING_DATE'), 
    a: "blank"
  }, {
    n: "shipping_by",
    r: ss.getRangeByName('Projects_SHIPPING_BY'), 
    a: "keep"
  }, {
    n: "address",
    r: ss.getRangeByName('Projects_ADDRESS'), 
    a: "formula",
    f: "=IF(ISBLANK(INDEX(Projects_STORE_ID,ROW(),0)),,VLOOKUP(INDEX(Projects_STORE_ID,ROW(),0),'Store Info'!$A:$O,12,FALSE))"
  }, {
    n: "notes",
    r: ss.getRangeByName('Projects_NOTES'), 
    a: "blank" 
  }, {
    n: "trello_card_id",
    r: ss.getRangeByName('Projects_TRELLO_CARD_ID'),
    a: "keep"
  }, {
    n: "trello_card_name",
    r: ss.getRangeByName('Projects_TRELLO_CARD_NAME'),
    a: "formula",
    f: "=IF(AND(NOT(ISBLANK(INDEX(Projects_STORE_ID,ROW(),0))),NOT(ISBLANK(INDEX(Projects_STORE_NAME,ROW(),0))),NOT(ISBLANK(INDEX(Projects_DESCRIPTION,ROW(),0)))),INDEX(Projects_STORE_ID,ROW(),0)&\" \"&INDEX(Projects_STORE_NAME,ROW(),0)&\" - \"&INDEX(Projects_DESCRIPTION,ROW(),0),)"
  }, {
    n: "row_timestamp",
    r: ss.getRangeByName('Projects_ROW_TIMESTAMP'),
    a: "code"
  }
];
