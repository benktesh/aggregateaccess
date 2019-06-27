This application can be used to combine tables from multiple database into one access database file.
The input.json file provides file paths for input databases, output database and tables.
If the output database does not exists, a new database file is created.

The content of input.json is as below:
 {
	 "combinedOutputFileName": "C:\\temp\\SIG\\FinalOut\\PC286_FVS_output.accdb",
	  "tableNames": [
		"FVS_Cases",
		"FVS_Carbon",
		"FVS_Hrv_Carbon",
		"FVS_Summary"
	  ],
	  "dataBases": [
		"C:\\temp\\SIG\\M0-Out\\PC286_FVS_output.accdb",
		"C:\\temp\\SIG\\M1-Out\\PC286_FVS_output.accdb",
		"C:\\temp\\SIG\\M2-Out\\PC286_FVS_output.accdb",
		"C:\\temp\\SIG\\M3-Out\\PC286_FVS_output.accdb"
	  ]
  }