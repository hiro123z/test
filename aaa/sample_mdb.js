var database;

var group1;
var group2;

onload = init;
onunload = dbClose;

function init() {


  dbConnect();
  //dataDisp();
  dispGroup1();
  //dbClose();
//fa();

}

//データベースに接続
function dbConnect() {
  database = new ActiveXObject("ADODB.Connection");
  database.Open("Driver={Microsoft Access Driver (*.mdb)}; DBQ=c:\\aaa\\s3.mdb;");
  //database.Open("Driver={Microsoft Access Driver (*.mdb)}; DBQ=.\\s3.mdb;");
  //alert("データベースに接続しました。");
}

//データベースを切断
function dbClose() {
  database.Close();
  database = null;
  alert("データベースを切断しました。");
}


//エンジン番号
function dispGroup1() {
    var mySql = "SELECT T01Prefecture.[グループ1] FROM T01Prefecture GROUP BY T01Prefecture.[グループ1];";
    var recordSet = database.Execute(mySql);
    var tempHtml="";
    document.getElementById("dispEngine").innerHTML = "";
	document.getElementById("dispDate").innerHTML = "";
	document.getElementById("dispFile").innerHTML = "";
    while (!recordSet.EOF){
		tempHtml = tempHtml + "<a class=\"engine\">" + recordSet(0) + "</a>" + "<br>";
		recordSet.MoveNext();
    }
    document.getElementById("dispEngine").innerHTML = tempHtml;
    recordSet.Close();
    recordSet = null;

	var obj = document.getElementsByClassName('engine');
	for (i = 0; i < obj.length; i++) {
		var obj2 = document.getElementsByClassName('engine');

		obj[i].addEventListener("click", function() {
			for (j = 0; j < obj2.length; j++) {
				obj2[j].classList.remove('red');
			}	
			this.classList.toggle('red');
			group1= this.innerHTML;
			dispGroup2();
		});
  }

}

//日付
function dispGroup2() {

    var mySql = "SELECT T01Prefecture.[グループ2] FROM T01Prefecture GROUP BY T01Prefecture.[グループ2], T01Prefecture.[グループ1] HAVING (((T01Prefecture.[グループ1])=" + group1 + "));";
    var recordSet = database.Execute(mySql);
    var tempHtml="";
    document.getElementById("dispDate").innerHTML = "";
	document.getElementById("dispFile").innerHTML = "";
	
    while (!recordSet.EOF){
		tempHtml = tempHtml + "<a class=\"date\">" + recordSet(0) + "</a>" + "<br>";
		recordSet.MoveNext();
    }
    document.getElementById("dispDate").innerHTML = tempHtml;
    recordSet.Close();
    recordSet = null;

	var obj = document.getElementsByClassName('date');
	for (i = 0; i < obj.length; i++) {
		obj[i].addEventListener("click", function() {
		var obj2 = document.getElementsByClassName('date');
		for (j = 0; j < obj2.length; j++) {
			obj2[j].classList.remove('red');
		}
			this.classList.toggle('red');
			group2= this.innerHTML;
			dispGroup3();
		});
  }
}

//ファイル名
function dispGroup3() {
    var mySql = "SELECT T01Prefecture.[グループ3] FROM T01Prefecture WHERE (((T01Prefecture.[グループ1])=" + group1 + ") AND ((T01Prefecture.[グループ2])=" + group2 + "));";
    var recordSet = database.Execute(mySql);
    var tempHtml="";
    document.getElementById("dispFile").innerHTML = "";
	
    while (!recordSet.EOF){
		tempHtml = tempHtml + "<a class=\"file\">" + recordSet(0) + "</a>" + "<br>";
		recordSet.MoveNext();
    }
    document.getElementById("dispFile").innerHTML = tempHtml;
    recordSet.Close();
    recordSet = null;

	var obj = document.getElementsByClassName('file');
	var sh = new ActiveXObject( "WScript.Shell" );
	for (i = 0; i < obj.length; i++) {
		obj[i].addEventListener("click", function() {
		sh.Run(this.innerHTML);
    });
  }
}





/*

function fa() {

//SELECT T01Prefecture.[ファイル名], T01Prefecture.[グループ] FROM T01Prefecture WHERE (((T01Prefecture.[グループ])=2000));
    var mySql = "SELECT T01Prefecture.[ファイル名], T01Prefecture.[グループ] FROM T01Prefecture WHERE (((T01Prefecture.[グループ])=2000));";
    var recordSet = database.Execute(mySql);

    var tempHtml="";
    document.getElementById("disp").innerHTML = "";
var count; count=0
    while (!recordSet.EOF){

      //tempHtml = tempHtml + "<a id=\"" + count + "\" target=\"_blank\" href=\"" + recordSet(0) + "\">" + recordSet(0) + "</a>"  + "<br>";
  tempHtml = tempHtml + "<a>" + recordSet(0) + "</a>" + "<br>";
      recordSet.MoveNext();
	  count=count+1;
    }
    document.getElementById("disp").innerHTML = tempHtml;
     recordSet.Close();
    recordSet = null;
  alert(count);
}
function fb() {
  var Atag = document.getElementsByTagName('a');//step1
  var sh = new ActiveXObject( "WScript.Shell" );
  for (i = 0; i < Atag.length; i++) {//step2
    Atag[i].addEventListener("click", function() {
		//sh.Run( "C:/aaa/tx.txt");
		sh.Run(this.innerHTML);
    });
  }
}

function fc() {
var sh = new ActiveXObject( "WScript.Shell" );
sh.Run( "C:/aaa/tx.txt");
}

*/




//データ表示
function dataDisp() {

    var mySql = "select * from T01Prefecture order by PREF_CD";
    var recordSet = database.Execute(mySql);

    var tempHtml="";
    document.getElementById("disp").innerHTML = "";
    while (!recordSet.EOF){
      tempHtml = tempHtml + recordSet(0) + ":" + recordSet(1) +":" +recordSet(2)+"<br>";
      recordSet.MoveNext();
    }
    document.getElementById("disp").innerHTML = tempHtml;
     recordSet.Close();
    recordSet = null;
/**/
}