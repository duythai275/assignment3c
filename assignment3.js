const XLSX = require('xlsx');
const fetch = require( "node-fetch" );
const access = require( "./access.json" );
const config = {
	sheets: [
		{
			sheetName: "dataElements",
			url: "http://dhis.academy/lao_25/api/dataElements.json?fields=:identifiable&paging=false",
			data: {}
		},
		{
			sheetName: "organisationUnits",
			url: "http://dhis.academy/lao_25/api/organisationUnits.json?fields=:identifiable&paging=false",
			data: {}
		},
		{
			sheetName: "indicators",
			url: "http://dhis.academy/lao_25/api/indicators.json?fields=:identifiable&paging=false",
			data: {}
		}
	]
};

const createAuthenticationHeader = (username, password) => {
  return "Basic " + new Buffer( username + ":" + password ).toString( "base64" );
};






  
  
  
  
  
/* set up workbook objects -- some of these will not be required in the future */
let wb = {}
wb.Sheets = {};
wb.SheetNames = [];


const createSheet = (data,sheetName) => {
	console.log(data);
	if(sheetName=="dataElements") data = data.dataElements;
	if(sheetName=="organisationUnits") data = data.organisationUnits;
	if(sheetName=="indicators") data = data.indicators;
	const ws_name = sheetName;
  
	// create worksheet:
	var ws = {}

	// the range object is used to keep track of the range of the sheet
	var range = {s: {c:0, r:0}, e: {c:100, r:0 }};
	
	data.forEach( (row,index) => {
		
		Object.keys( row )
		//.forEach( k => wjf.sync( regions[k].filename, regions[k].data ) );
		.forEach( (k,i) => {
			const cell = { v: row[k] };
			const cell_ref = XLSX.utils.encode_cell({c:i,r:index});
			
			ws[cell_ref] = cell;
		} );
		
		range.e.r = index;
		
	} );
	
	ws['!ref'] = XLSX.utils.encode_range(range);
	//add worksheet to workbook
	wb.SheetNames.push(ws_name);
	wb.Sheets[ws_name] = ws;

	
};


const loadJson = i => {
	console.log(i);
	if(i==config.sheets.length){
		config.sheets.forEach( sheet => {
			createSheet(sheet.data,sheet.sheetName);
		});
		//write file
		XLSX.writeFile(wb, 'assignment3c.xlsx');
	}
}

let i = 0;
config.sheets.forEach( sheet => {
	fetch(
		sheet.url,
		{
			headers: {
				Authorization: createAuthenticationHeader( access.username, access.password )
			}
		}
	)
	.then( result => result.json() )
	.then( data => {
		sheet.data = data;
		i++;
		//createSheet( data, sheet.sheetName );
		loadJson(i);
	} );
});


