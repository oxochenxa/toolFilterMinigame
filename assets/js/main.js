/* oss.sheetjs.com (C) 2014-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */

/** drop target **/
var _target = document.getElementById('drop');
var _file = document.getElementById('file');
var _grid = document.getElementById('grid');
var _gobalData;
var _deleteData = new Array();
var sheetname;
/** Spinner **/
var spinner;

var _workstart = function() { spinner = new Spinner().spin(_target); }
var _workend = function() { spinner.stop(); }

/** Alerts **/
var _badfile = function() {
  alertify.alert('This file does not appear to be a valid Excel file.  If we made a mistake, please send this file to <a href="mailto:dev@sheetjs.com?subject=I+broke+your+stuff">dev@sheetjs.com</a> so we can take a look.', function(){});
};

var _pending = function() {
  alertify.alert('Please wait until the current file is processed.', function(){});
};

var _large = function(len, cb) {
  alertify.confirm("This file is " + len + " bytes and may take a few moments.  Your browser may lock up during this process.  Shall we play?", cb);
};

var _failed = function(e) {
  console.log(e, e.stack);
  alertify.alert('We unfortunately dropped the ball here.  Please test the file using the <a href="/js-xlsx/">raw parser</a>.  If there are issues with the file processor, please send this file to <a href="mailto:dev@sheetjs.com?subject=I+broke+your+stuff">dev@sheetjs.com</a> so we can make things right.', function(){});
};

/* make the buttons for the sheets */
var make_buttons = function(sheetnames, cb) {
  var buttons = document.getElementById('buttons');
  buttons.innerHTML = "";
  sheetname = sheetnames;
  sheetnames.forEach(function(s,idx) {
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.name = 'btn' + idx;
    btn.text = s;
    var txt = document.createElement('h3'); txt.innerText = s; btn.appendChild(txt);
    btn.addEventListener('click', function() { cb(idx); }, false);
    buttons.appendChild(btn);
    buttons.appendChild(document.createElement('br'));
  });
};

var cdg = canvasDatagrid({
  parentNode: _grid
});
cdg.style.height = '100%';
cdg.style.width = '100%';

function _resize() {
  _grid.style.height = (window.innerHeight - 200) + "px";
  _grid.style.width = (window.innerWidth - 200) + "px";
}
window.addEventListener('resize', _resize);

var _onsheet = function(json, sheetnames, select_sheet_cb) {

  make_buttons(sheetnames, select_sheet_cb);

  /* show grid */
  _grid.style.display = "block";
  _resize();

  /* set up table headers */
  var L = 0;
  json.forEach(function(r) { if(L < r.length) L = r.length; });
  for(var i = json[0].length; i < L; ++i) {
    json[0][i] = "";
  }

  /* load data */
  cdg.data = json;

  /* remove array null */
  cdg.data = removeBlankData(cdg.data);
  _gobalData = cdg.data;
};

/** Drop it like it's hot **/
DropSheet({
  file: _file,
  drop: _target,
  on: {
    workstart: _workstart,
    workend: _workend,
    sheet: _onsheet,
    foo: 'bar'
  },
  errors: {
    badfile: _badfile,
    pending: _pending,
    failed: _failed,
    large: _large,
    foo: 'bar'
  }
})

/** Filter Date **/
function filterdate() {
  var filterdate = new Date(document.getElementById("filterdate").value);
  if(!filterdate){
    alert("Bạn chưa nhập ngày mà đòi lọc ngày à!");
    $("#checkbox_filterdate").prop("checked", false);
  }else {
    cdg.data = _gobalData;
    for(i = 1; i <= cdg.data.length; i++){
      if(cdg.data[i]){
        var toDate = new Date(cdg.data[i][0]).getDate();
        var toMonth = new Date(cdg.data[i][0]).getMonth()+1;
        var toYear = new Date(cdg.data[i][0]).getFullYear();
        var originalDate = new Date(toYear +'-'+ toMonth +'-'+ toDate);
      }

      if (isLater(filterdate, originalDate)){
        if(cdg.data[i]) {
          cdg.data[i].push("Lỗi Ngày");
        }
        _deleteData.push(cdg.data[i]);
        cdg.data.splice(i, 1);
      }
    }
  }
}

function filterphone() {
  for(i = 1; i < cdg.data.length - 1; i++){
    var j = i + 1;
    if(cdg.data[i]){
      if(cdg.data[i][4].length > 10 | cdg.data[i][4].length < 9){
        cdg.data[i].push("Lỗi SĐT");
        _deleteData.push(cdg.data[i]);
        cdg.data.splice(i, 1);
      }else {
        while (j < cdg.data.length){
          if (cdg.data[3359])
          if(cdg.data[i][4] == cdg.data[j][4] & i != j){
            cdg.data[j].push("Lỗi SĐT");
            _deleteData.push(cdg.data[j]);
            cdg.data.splice(j, 1);
          }
          else j += 1;
        }
      }
    }
  }
}

function filtercmnd() {
  for(i = 1; i <= cdg.data.length; i++){
    if(cdg.data[i]){
      if(cdg.data[i][2].length > 3){
        if(cdg.data[i]) {
          cdg.data[i].push("Lỗi CMND");
        }
        _deleteData.push(cdg.data[i]);
        cdg.data.splice(i, 1);
      }
    }
  }
}
function show_data_remove() {
  if (cdg.data == null || _deleteData == null)
    alert("Làm gì có data mà đòi hiện!!!")
  else{
    var str = "";
    str += '<thead><tr>';
    for(i = 0; i < cdg.data[0].length; i++){
      if (cdg.data[0][i])
        str += '<th>'+cdg.data[0][i]+'</th>';
    }
    str += '<th>Lỗi</th>';
    str += '</tr></thead>';
    str += '<tbody>';

    for(i = 0; i < _deleteData.length; i++){
      str += '<tr>';
      for(j = 0; j < _deleteData[i].length; j++){
        if (_deleteData[i][j])
          str += '<th>'+_deleteData[i][j]+'</th>';
      }
      str += '</tr>';
    }
    str += '</tbody>';
    $('table#remove_data').append(str);
    $('#remove_data').DataTable(
        {
          pageLength : 5
        }
    );
  }
}
function isLater(filterdate, originalDate) {
  return filterdate < originalDate;
}

function removeBlankData(array) {
  array = array.filter(item => item.length != 0);
  return array;
};

function printFileAfterFilter(){
  if (cdg.data == null)
    alert("làm gì có data mà xuất!!!")
  else{
    /* Sheet Name */
    var ws_name = String(sheetname);
    var fileNameAfterFilter = "Danh sách sau khi lọc.xlsx";
    var wb = XLSX.utils.book_new(),
        ws = XLSX.utils.aoa_to_sheet(cdg.data);

    /* Add worksheet to workbook */
    XLSX.utils.book_append_sheet(wb, ws, ws_name);

    /* bookType can be 'xlsx' or 'xlsm' or 'xlsb' */
    var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };

    var wbout = XLSX.write(wb,wopts);

    function s2ab(s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }

    /* the saveAs call downloads a file on the local machine */
    saveAs(new Blob([s2ab(wbout)],{type:""}), fileNameAfterFilter);
  }
}

function printFileRemove(){
  if (_deleteData.length == 0)
    alert("làm gì có data mà xuất!!!")
  else {
    /* Sheet Name */
    var ws_name = String(sheetname);
    var fileNameAfterFilter = "Danh sách xoá.xlsx";
    var wb = XLSX.utils.book_new(),
        ws = XLSX.utils.aoa_to_sheet(_deleteData);

    /* Add worksheet to workbook */
    XLSX.utils.book_append_sheet(wb, ws, ws_name);

    /* bookType can be 'xlsx' or 'xlsm' or 'xlsb' */
    var wopts = {bookType: 'xlsx', bookSST: false, type: 'binary'};

    var wbout = XLSX.write(wb, wopts);

    function s2ab(s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }

    /* the saveAs call downloads a file on the local machine */
    saveAs(new Blob([s2ab(wbout)], {type: ""}), fileNameAfterFilter);
  }
}