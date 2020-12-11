Set-ExecutionPolicy RemoteSigned

Add-Type –Path ‘C:\Program Files (x86)\MySQL\MySQL Connector Net 8.0.22\Assemblies\v4.5.2\MySql.Data.dll'
$Connection = [MySql.Data.MySqlClient.MySqlConnection]@{ConnectionString='server=192.168.3.228;uid=aduser;pwd=wildbelwest;database=sap'}
$Connection.Open()
$MYSQLCommand = New-Object MySql.Data.MySqlClient.MySqlCommand
$MYSQLDataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter
$MYSQLDataSet = New-Object System.Data.DataSet
$MYSQLCommand.Connection=$Connection
$MYSQLCommand.CommandText='SELECT mobil, work, home, timesheet FROM contacts;'
$MYSQLDataAdapter.SelectCommand=$MYSQLCommand
$NumberOfDataSets=$MYSQLDataAdapter.Fill($MYSQLDataSet, "data")

Import-Module activedirectory
$BWADPath = 'OU=ÑÎÎÎ Áåëâåñò,DC=belwest,DC=corp'

foreach($DataSet in $MYSQLDataSet.tables[0])
{	
	$tabnum = $DataSet.timesheet
	$user = get-ADUser -Searchbase $BWADPath -Filter {EmployeeID -eq $tabnum} -Properties telephoneNumber
	If ($null -ne $user ) 
	{
		Set-ADUser $user -Replace @{telephoneNumber = $DataSet.home}
	}
}

$Connection.Close()