var connection = new ActiveXObject("ADODB.Connection") ;

var connectionstring="Data Source=efhsql;Initial Catalog=hblive;User ID=reports2;Password=;Provider=SQLOLEDB";

connection.Open(connectionstring);
var rs = new ActiveXObject("ADODB.Recordset");

rs.Open("SELECT * FROM employeebirthdays", connection);
rs.MoveFirst
while(!rs.eof)
{
   document.write(rs.fields(1));
   rs.movenext;
}

rs.close;
connection.close;
