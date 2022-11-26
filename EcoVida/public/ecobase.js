var connection = new ActiveXObject("ADODB.Connection") ;

var connectionstring="Data Source=servidorecovida.database.windows.net;Initial Catalog=ecobase;User ID=enoc.lopez;Password=jerez1405:3;Connect Timeout=60;Encrypt=True;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

connection.Open(connectionstring);
var rs = new ActiveXObject("ADODB.Recordset");

rs.Open("Usuario", connection);
rs.MoveFirst
while(!rs.eof)
{
   document.write(rs.fields(1));
   rs.movenext;
}

rs.close;
connection.close; 