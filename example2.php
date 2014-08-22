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


$con=mysqli_connect("localhost","root","","sddr");
// Check connection
if (mysqli_connect_errno()) {
  echo "Failed to connect to MySQL: " . mysqli_connect_error();
}

$letters = range('A', 'Z'); // column letters
$sheet_index = 3;
$table_name = "data_structure";
$data_structure_type = "Organisation Data Structure";
$category_id = "6";
$sub_category_id = "3";

$i = 2;

for ($i; $i <= $data->rowcount($sheet_index); $i++) {

	if ($i != $i-1) $insert = "";

	for ($j = 0; $j < $data->colcount($sheet_index); $j++) {

		if ($j != $data->colcount($sheet_index)-1) {
			$insert .= "'".mysqli_real_escape_string($con,$data->val($i,$letters[$j], $sheet_index))."', ";
		} else {
			$insert .= "'".mysqli_real_escape_string($con,$data->val($i,$letters[$j], $sheet_index))."'";
		}
	}

	echo "<br/>INSERT INTO `".$table_name."`(data_structure_type, element_name, element_description, rational, field_namecode, alias, field_size, element_owner, data_sub_category_id, category_id) VALUES('".$data_structure_type."', ".$insert.", '".$sub_category_id."', '".$category_id."');<br/>";
}



// Execute query
/*if (mysqli_query($con,$sql)) {
  echo "Table persons created successfully";
} else {
  echo "Error creating table: " . mysqli_error($con);
}*/

 ?>
</body>
</html>
