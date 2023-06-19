var database;

onload = init;
onunload = dbClose;

function init() {
  document.getElementById("txtPrefCd").onblur = function (){blur(this);}
  document.getElementById("txtPrefCd").onfocus = function (){focus(this);}
  document.getElementById("txtPrefName").onblur = function (){blur(this);}
  document.getElementById("txtPrefName").onfocus = function (){focus(this);}

  dbConnect();
}

//データベースに接続
function dbConnect() {
  database = new ActiveXObject("ADODB.Connection");
  database.Open("Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=c:\\aaa\\s3.mdb;");
  alert("データベースに接続しました。");
}

//データベースを切断
function dbClose() {
  database.Close();
  database = null;
  alert("aデータベースを切断しました。");
}

function focus(obj){
  obj.style.backgroundColor = "#ffff00";
}

function blur(obj){
  obj.style.backgroundColor = "#ffffff";
}