// Pull in config file
var sco = new ActiveXObject("Scripting.FileSystemObject");
var configfile = new ActiveXObject("MSXML2.DOMDocument");
configfile.async=false;

var workingdir = sco.GetParentFolderName(WScript.ScriptFullName); // Overcome working directory problems
var basename = sco.GetBaseName(WScript.ScriptFullName);
var configfilename = new String(basename + ".xml");
var outputfilename = new String(basename + "-output.html");

//configfile.load(workingdir + "\\config.xml");
configfile.load(workingdir + "\\" + configfilename);

var adgroupnodes = configfile.getElementsByTagName("adgroup");
var adgroups = new Array();
var idstring = new String(";");

while (node = adgroupnodes.nextNode)
	{
	var attlist = node.attributes;
	adgroups.push(attlist.item(0).value);
	}
// Reprovision XML object
var htmlnode = configfile.createNode(1, "html", "http://www.w3.org/1999/xhtml");
var headnode = configfile.createElement("head");
var bodynode = configfile.createElement("body");
htmlnode.appendChild(headnode);
htmlnode.appendChild(bodynode);
headnode.appendChild(makeHTMLNode("title", "HRList to AD Group Comparison"));
bodynode.appendChild(makeHTMLNode("h1", "HRList to AD Group Comparison"));

configfile.replaceChild(htmlnode, configfile.documentElement);

// Open Excel File

var exxon = new ActiveXObject("Excel.Application");
exxon.visible = true;
var exWbks = exxon.Workbooks;

var xlWBATWorksheet = -4167;
var xlDown = -4121;

// Populate idstring
hrfile = exWbks.Open(WScript.Arguments.Item(0), 0, true );
with (hrfile)
	{
	Worksheets(1).Range("A2").Select;
	var justarange = Worksheets(1).Range(exxon.Selection, exxon.Selection.End(xlDown));
	var jsarray = justarange.Value.toArray();
	idstring = idstring.concat(jsarray.join(";") + ";");
	}
hrfile.Close(false);
exxon.Quit();

// Connect to AD and do stuff

var gcroot = new String("<GC://dc=avonet,dc=net>;");
var adoconn = new ActiveXObject("ADODB.Connection");
var adocomm = new ActiveXObject("ADODB.Command");

adoconn.Provider = ("ADSDSOObject");
adoconn.Open;

adocomm.ActiveConnection = adoconn;

// Variables for main loop
var matchedaccounts = new Array(); // These need to be reported
var members = new Array(); // Userlist in the group
var memberof = new Array(); // Grouplist in the user
var disableflag = new Array(); // Will store boolean value for whether or not the user is disabled.
//var cachedConn = GetObject("LDAP://dc=avonet, dc=net");

while (adgroups.length > 0) //Iterate through each group
	{
	var adgroupname = adgroups.pop();

	//WScript.Echo(adgroupname);
	var filter = "(&(objectclass=group)(cn=" + adgroupname + "))";
	adocomm.CommandText=gcroot + filter +";" + "member;subtree";
	var records = adocomm.Execute;

	//WScript.Echo(typeof records.Fields("member").Value);
	//WScript.Echo(records.Fields("member").Value.toArray().length);
	while (records.EOF == false && (typeof records.Fields("member").Value == "unknown"))
	//while (records.EOF == false)
		{	
		members = records.Fields("member").Value.toArray();
		//WScript.Echo("blah");

		while (members.length > 0)
			{
			var userlookup = GetObject("GC://" + members.pop());
			
			if (typeof userlookup == "object")
				{
				var eidWithAnchors = new String(";" + userlookup.EmployeeID + ";");
				var eidpattern = new RegExp(eidWithAnchors);

				if (eidpattern.test(idstring) == true)
					{
					var samname = new String(userlookup.Get("sAMAccountName"));

					if (typeof matchedaccounts[samname] == "object")
						{
						matchedaccounts[samname].push(adgroupname);
						}
					else
						{
						matchedaccounts[samname] = new Array(adgroupname);
						}

					if (!disableflag[samname]) {disableflag[samname] = userlookup.AccountDisabled} // See if disabled
					}

				}
				
			}

		records.MoveNext;
		}

	}

// Spit out HTML
for (matchedaccount in matchedaccounts)
	{
	var ul = makeAnUnorderedList(matchedaccounts[matchedaccount]);
	if (disableflag[matchedaccount]) { matchedaccount = matchedaccount.concat(" (DISABLED)") }
	var divnode = makeHTMLNode("div", matchedaccount);
	divnode.appendChild(ul);
	bodynode.appendChild(divnode);
	}

if (bodynode.childNodes.length <= 1)
	{
	bodynode.appendChild(makeHTMLNode("h2", "No Matches"));
	}

configfile.save(workingdir + "\\" + outputfilename);
WScript.Echo("Wrote " + outputfilename);

// FUNCTIONS
function makeAnUnorderedList(arrayofelements)
	{
	var ul = configfile.createElement("ul");
	for (i in arrayofelements)
		{
		ul.appendChild(makeHTMLNode("li", arrayofelements[i]));
		}
	return ul;
	}

function makeHTMLNode(nodename, content)
	{
	var node = configfile.createElement(nodename);
	node.appendChild(configfile.createTextNode(content));
	return node;
	}
