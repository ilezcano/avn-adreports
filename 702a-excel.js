/* Script by Ismael Lezcano
 * for control 702.A
 * Revision 20090422
 *
 * This script performs the following functions
 * 1. Using a Feed from HR by clicking and dragging the file over the script file,
 * the script opens the file in EXCEL. The employed column, which is leftmost in 
 * all of our HR feeds, is selected and concatonated onto a script object. After
 * the values are copied, the file is closed.
 * 2. A new EXCEL workbook is created with each report occupying its own sheet. Two
 * special reports have thier own sheets; ADDump and Not in HR File.
 * 3. Using ActiveX Data Objects, the script connects to Active Directory and retrieves
 * all non-disabled user objects in the OUs for which the control regulates. As the
 * records are retrieved, they are immediately evaluated and divided to the appropriate
 * report based on Employee ID.
 * 4. If the employee does not match the static criteria (Blank, consultant, service,
 * generic, etc.) it is evaluated against the data previously retrieved from the
 * HR file. If the employee ID is not found, it is reported to the sheet thusly titled.
 * 5. Regardless of the result of the above tests, the record is recorded into the ADDump
 * sheet for reference.
 */
// Objects
var exxon = new ActiveXObject("Excel.Application");
exxon.visible = true;
var exWbks = exxon.Workbooks;
var adoconn = new ActiveXObject("ADODB.Connection");
var adocomm = new ActiveXObject("ADODB.Command");
//
// Constants
//
var xlWBATWorksheet = -4167;
var xlDown = -4121;
// 
// Variables
//
var sheetheader = new Array("Last Name", "FirstName", "UserName", "Description" ,"Department", "Distinguished Name", "Manager", "Create Date", "userAccountControl");
var reportnames = new Array("generic", "service", "consultant", "blanks");
var idstring = new String(";");
var attarray = new Array("employeeid", "sn", "givenName", "sAMAccountName", "description", "department", "distinguishedName", "manager", "whenCreated", "userAccountControl");
var ouarray = new Array("extranet", "lax", "mia", "rye", "glb", "nyc", "suf", "corp", "us,ou=dsm");
var filter = "(&(objectclass=user) (useraccountcontrol:1.2.840.113556.1.4.803:=512)(!(useraccountcontrol:1.2.840.113556.1.4.803:=2)))";
var bases = new Array();
// 
// Build filters
// 
for (i in ouarray)
	{
		bases.push("<LDAP://ou=" + ouarray[i] + ",dc=na,dc=avonet,dc=net>;");
	}
//
// Start the Commotion
//
WScript.Echo("Please don't touch anything untill the message that says 'Done!', or any error message. This will take a while to cook.");
hrfile = exWbks.Open(WScript.Arguments.Item(0), 0, true );
with (hrfile)
	{
	Worksheets(1).Range("A2").Select;
	var justarange = Worksheets(1).Range(exxon.Selection, exxon.Selection.End(xlDown));
	/*
	for (i = 1; i <= justarange.Count; i++)
		{
		var cellvalue = justarange.Cells(i).Value + ";";
		idstring = idstring.concat(cellvalue);
		}
	*/
	var jsarray = justarange.Value.toArray();
	idstring = idstring.concat(jsarray.join(";") + ";");
	}
hrfile.Close(false);

var exWb = exWbks.Add(xlWBATWorksheet);
//
// Set up special reports
//
with (exWb)
	{
	Worksheets(1).Name = "ADDump";
	var adsheetheader = sheetheader.slice(0);
	adsheetheader.unshift("Employee ID");
	var headerrange = ActiveSheet.Range("A1", ActiveSheet.Cells(1, adsheetheader.length));
	headerrange.Font.Bold = true;
	ActiveSheet.ListObjects.Add(1, headerrange, false, 1);

	for (var anindex = 1; anindex <= adsheetheader.length; anindex++)
		{
		headerrange.Cells(anindex).Value = adsheetheader[anindex - 1];
		}
	
	Worksheets.Add();
	ActiveSheet.Name = "Not in HR File";
	
	headerrange = ActiveSheet.Range("A1", ActiveSheet.Cells(1, adsheetheader.length));
	ActiveSheet.ListObjects.Add(1, headerrange, false, 1);

	for (var anindex = 1; anindex <= adsheetheader.length; anindex++)
		{
		headerrange.Cells(anindex).Value = adsheetheader[anindex - 1];
		}
	}

//
// Set up other reports
//
for (i = 0; i < reportnames.length; i++)
	{
	exWb.Worksheets.Add();
	exWb.ActiveSheet.Name = reportnames[i];
	var headerlength = sheetheader.length;
	var headerrange = exWb.ActiveSheet.Range("A1", exWb.ActiveSheet.Cells(1,headerlength));
	headerrange.Font.Bold = true;

	for (anindex = 1; anindex <= headerlength; anindex++)
		{
		headerrange.Cells(anindex).Value = sheetheader[anindex - 1];
		}
	
	exWb.ActiveSheet.ListObjects.Add(1, headerrange, false ,1);
	}
exWb.WorkSheets("ADDump").Activate;

//
// Start Query
//
adoconn.Provider = ("ADSDSOObject");
adoconn.Open;

adocomm.ActiveConnection = adoconn;
for (ouindex in bases)
	{
	adocomm.CommandText=bases[ouindex] + filter +";" + attarray + ";subtree";
	adocomm.Properties("Page Size") = 1000;
	var records = adocomm.Execute;
	while (records.EOF == false)
		{
		var resultarray = new Array(); // Correspond to Columns
		var eid = new String(records.Fields("employeeid").ActualSize > 0 ? records.Fields("employeeid") : "");
		with (records)
			{
			resultarray[0] = Fields("sn").ActualSize > 0 ? Fields("sn") : "";
			resultarray[1] = Fields("givenName").ActualSize > 0 ? Fields("givenName") : "";
			resultarray[2] = Fields("sAMAccountName").ActualSize > 0 ? Fields("sAMAccountName") : "";
			resultarray[3] = Fields("description").ActualSize > 0 ? Fields("description") : "";
			resultarray[4] = Fields("department").ActualSize > 0 ? Fields("department") : "";
			resultarray[5] = Fields("distinguishedName").ActualSize > 0 ? Fields("distinguishedName") : "";
			resultarray[6] = Fields("manager").ActualSize > 0 ? Fields("manager") : "";
			resultarray[7] = Fields("whenCreated").ActualSize > 0 ? Fields("whenCreated") : "";
			resultarray[8] = Fields("userAccountControl").ActualSize > 0 ? Fields("userAccountControl") : "";
			}
		var adcolumns = resultarray.slice(0); // Correspond to AD and HRFile columns
		adcolumns.unshift(eid);
		if (eid.match(/^$/))
			{cookiereport("blanks", resultarray);}
		else if (eid.match(/^service$/i))
			{cookiereport("service", resultarray);}
		else if (eid.match(/^consultant$/i))
			{cookiereport("consultant", resultarray);}
		else if (eid.match(/^generic$/i))
			{cookiereport("generic", resultarray);}
		else if (eid.length > 0)
			{
			var re = new RegExp(";"+eid+";");
			if (re.test(idstring) == false)
				{
				cookiereport("Not in HR File", adcolumns);
				}
			}

		cookiereport("ADDump", adcolumns);
		records.MoveNext;
		}
	}

adoconn.close;
WScript.Echo("Done!");


function cookiereport(title , fields )
	{
	var newlistrow = exWb.WorkSheets(title).ListObjects(1).ListRows.Add();
	var newrange = newlistrow.Range;
	newrange.Font.Bold = false;
	for (var i = 0; i < fields.length; i++)
		{	
		newrange.Cells(i+1).Value=fields[i];
		}

	return;
	}
