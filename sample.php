<?php
require('XLSXReader.php');
$xlsx = new XLSXReader('sample.xlsx');
$sheetNames = $xlsx->getSheetNames();

?>
<!DOCTYPE html>
<html>
<head>
	<title>XLSXReader Sample</title>
	<style>
		body {
			font-family: Helvetica, sans-serif;
			font-size: 12px;
		}

		table, td {
			border: 1px solid #000;
			border-collapse: collapse;
		}
	</style>
</head>
<body>
<h1>XLSXReader Example</h1>

List of the sheets in this workbook (indexed by sheetId):<br>
<?=debug($sheetNames);?>

<?
foreach($sheetNames as $sheetName) {
	$sheet = $xlsx->getSheet($sheetName);
	?>
	<h1><?=escape($sheetName);?></h1>
	<?
	array2Table($sheet->getData());
}

?>
</body>
</html>




<?
function array2Table($data) {
	echo '<table>';
	foreach($data as $row) {
		echo "<tr>";
		foreach($row as $cell) {
			echo "<td>";
			echo escape($cell);
			echo "</td>";
		}
		echo "</tr>";
	}
	echo '</table>';
}

function debug($data) {
	echo '<pre>';
	print_r($data);
	echo '</pre>';
}

function escape($string) {
	return htmlspecialchars($string, ENT_QUOTES);
}
