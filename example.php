<?php
error_reporting(E_ALL ^ E_NOTICE);
require_once 'excel_reader2.php';
$data = new Spreadsheet_Excel_Reader("Data-convert.xls");
?>
<html>
<head>
<style>
table.excel {
	border-style:ridge;
	border-width:1;
	border-collapse:collapse;
	font-family:sans-serif;
	font-size:12px;
}
table.excel thead th, table.excel tbody th {
	background:#CCCCCC;
	border-style:ridge;
	border-width:1;
	text-align: center;
	vertical-align:bottom;
}
table.excel tbody th {
	text-align:center;
	width:20px;
}
table.excel tbody td {
	vertical-align:bottom;
}
table.excel tbody td {
    padding: 0 3px;
	border: 1px solid #EEEEEE;
}
</style>
</head>

<body>
<?php


$con=mysqli_connect("localhost","root","","sddr_data");
// Check connection
if (mysqli_connect_errno()) {
  echo "Failed to connect to MySQL: " . mysqli_connect_error();
}

$letters = range('A', 'Z'); // column letters
$sheet_index = 4;
$table_name = "code_list_lookup";

// echo $data->dump(false, false, $sheet_index);

// loop through sheets
for ($i = 0; $i < $data->colcount($sheet_index); $i++) {
	// concatenate strings
	$create .= "`".strtolower(
				str_replace('/', '',
					str_replace(' ', '_', $data->val(1,$letters[$i],$sheet_index))
				)
		)."` varchar(100) NOT NULL,";
}

// Create table
$createQuery = "CREATE TABLE ".$table_name."(
	`id` int(10) unsigned NOT NULL AUTO_INCREMENT,".$create."
	PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;<br/>";

echo $createQuery;

$i = 2;

for ($i; $i <= $data->rowcount($sheet_index); $i++) {
	// concatenate strings
	// $e .= "`".strtolower(
	// 			str_replace('/', '',
	// 				str_replace(' ', '_', $data->val($i,$letters[$i],$sheet_index))
	// 			)
	// 	)."` varchar(100) NOT NULL,";

	if ($i != $i-1) $insert = "";

	for ($j = 0; $j < $data->colcount($sheet_index); $j++) {

		if ($j != $data->colcount($sheet_index)-1) {
			$insert .= "'".mysqli_real_escape_string($con,$data->val($i,$letters[$j], $sheet_index))."', ";
		} else {
			$insert .= "'".mysqli_real_escape_string($con,$data->val($i,$letters[$j], $sheet_index))."'";
		}
	}

	echo "<br/>INSERT INTO `".$table_name."` VALUES('".$i."', ".$insert.");<br/>";
}



echo $insertQuery;


// Execute query
/*if (mysqli_query($con,$sql)) {
  echo "Table persons created successfully";
} else {
  echo "Error creating table: " . mysqli_error($con);
}*/

 ?>
</body>
</html>
