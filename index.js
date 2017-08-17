const XLSX = require('xlsx');
const fetch = require("node-fetch");
const access = require("./access.json");

const createAuthenticationHeader = (username, password) => {
	return "Basic " + new Buffer(username + ":" + password).toString("base64");
};

const fetchFromDhis2 = url => {
	const headers = {
		Authorization: createAuthenticationHeader(access.username, access.password)
	};

	return fetch(url, {
		headers
	}).then(result => result.json());
};



const config = {
	sheets: [{
		sheetName: "dataElements",
		url: "https://hmis.gov.la/api/dataElements.json?fields=:identifiable&paging=false"
	}, {
		sheetName: "organisationUnits",
		url: "https://hmis.gov.la/api/organisationUnits.json?fields=:identifiable&paging=false"
	}, {
		sheetName: "indicators",
		url: "https://hmis.gov.la/api/indicators.json?fields=:identifiable&paging=false"
	}]
};



/* set up workbook objects -- some of these will not be required in the future */
let wb = {}
wb.Sheets = {};
wb.SheetNames = [];


const createSheet = (data, sheetName) => {
	const ws_name = sheetName;

	// create worksheet:
	var ws = {}

	// the range object is used to keep track of the range of the sheet
	var range = {
		s: {
			c: 0,
			r: 0
		},
		e: {
			c: 100,
			r: 0
		}
	};

	data[sheetName].forEach((row, index) => {

		Object.keys(row)
			//.forEach( k => wjf.sync( regions[k].filename, regions[k].data ) );
			.forEach((k, i) => {
				const cell = {
					v: row[k]
				};
				const cell_ref = XLSX.utils.encode_cell({
					c: i,
					r: index
				});

				ws[cell_ref] = cell;
			});

		range.e.r = index;

	});

	ws['!ref'] = XLSX.utils.encode_range(range);
	//add worksheet to workbook
	wb.SheetNames.push(ws_name);
	wb.Sheets[ws_name] = ws;


};



Promise.all(
	config.sheets.map(sheet => {
		return fetchFromDhis2(sheet.url);
	})
).then(arr => {

	config.sheets.forEach((sheet, i) => {
		createSheet(arr[i], sheet.sheetName);
	});

	XLSX.writeFile(wb, 'assignment3c.xlsx');

});