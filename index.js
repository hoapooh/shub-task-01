const axios = require("axios");
const xlsx = require("xlsx");
const fs = require("fs");

const BASE_URL = "https://go.microsoft.com/fwlink/?LinkID=521962";

(async () => {
	try {
		// 1: Fetch data from link
		const response = await axios.get(BASE_URL, { responseType: "arraybuffer" });

		// 2: Read the Excel file into a workbook
		const workbook = xlsx.read(response.data, { type: "buffer" });

		// Get the first sheet in the workbook
		const worksheet = workbook.Sheets[workbook.SheetNames[0]];

		// Convert the sheet to JSON for easier manipulation
		const data = xlsx.utils.sheet_to_json(worksheet);

		// 3: Filter rows where '  Sales ' > 50,000
		const filteredData = data.filter((row) => {
			const salesValue = Number(parseFloat(row["  Sales "]));
			return salesValue > 50000;
		});

		// 4: Create a new worksheet with the filtered data
		const newWorkbook = xlsx.utils.book_new();
		const newWorksheet = xlsx.utils.json_to_sheet(filteredData);
		xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "FilteredData");

		// 5: Write the new workbook to a file
		// This file will be created in the same directory as the script
		const outputFilePath = "filtered_data.xlsx";
		xlsx.writeFile(newWorkbook, outputFilePath);
	} catch (error) {
		if (error.code === "ENOTFOUND") {
			console.error("Network error: Please check your internet connection.");
		} else if (error.response) {
			console.error(
				`Server responded with status code ${error.response.status}: ${error.response.statusText}`
			);
		} else if (error.request) {
			console.error("No response received from the server.");
		} else {
			console.error("An error occurred:", error.message);
		}
	}
})();
