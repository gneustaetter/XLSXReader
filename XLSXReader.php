<?php
/*
XLSXReader
Greg Neustaetter <gneustaetter@gmail.com>
Artistic License

XLSXReader is a heavily modified version of:
	SimpleXLSX php class v0.4 (Artistic License)
	Created by Sergey Schuchkin from http://www.sibvision.ru - professional php developers team 2010-2011
	Downloadable here: http://www.phpclasses.org/package/6279-PHP-Parse-and-retrieve-data-from-Excel-XLS-files.html

Key Changes include:
	Separation into two classes - one for the Workbook and one for Worksheets
	Access to sheets by name or sheet id
	Use of ZIP extension
	On-demand access of files inside zip
	On-demand access to sheet data
	No storage of XML objects or XML text
	When parsing rows, include empty rows and null cells so that data array has same number of elements for each row
	Configuration option for removing trailing empty rows
	Better handling of cells with style information but no value
	Change of class names and method names
	Removed rowsEx functionality including extraction of hyperlinks
*/

class XLSXReader {
	protected $sheets = array();
	protected $sharedstrings = array();
	protected $sheetInfo;
	protected $zip;
	public $config = array(
		'removeTrailingRows' => true
	);
	
	// XML schemas
	const SCHEMA_OFFICEDOCUMENT  =  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
	const SCHEMA_RELATIONSHIP  =  'http://schemas.openxmlformats.org/package/2006/relationships';
	const SCHEMA_OFFICEDOCUMENT_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	const SCHEMA_SHAREDSTRINGS =  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
	const SCHEMA_WORKSHEETRELATION =  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

	public function __construct($filePath, $config = array()) {
		$this->config = array_merge($this->config, $config);
		$this->zip = new ZipArchive();
		$status = $this->zip->open($filePath);
		if($status === true) {
			$this->parse();
		} else {
			throw new Exception("Failed to open $filePath with zip error code: $status");
		}
	}

	// get a file from the zip
	protected function getEntryData($name) {
		$data = $this->zip->getFromName($name);
		if($data === false) {
			throw new Exception("File $name does not exist in the Excel file");
		} else {
			return $data;
		}
	}

	// extract the shared string and the list of sheets
	protected function parse() {
		$sheets = array();
		$relationshipsXML = simplexml_load_string($this->getEntryData("_rels/.rels"));
		foreach($relationshipsXML->Relationship as $rel) {
			if($rel['Type'] == self::SCHEMA_OFFICEDOCUMENT) {
				$workbookDir = dirname($rel['Target']) . '/';
				$workbookXML = simplexml_load_string($this->getEntryData($rel['Target']));
				foreach($workbookXML->sheets->sheet as $sheet) {				
					$r = $sheet->attributes('r', true);
					$sheets[(string)$r->id] = array(
						'sheetId' => (int)$sheet['sheetId'],
						'name' => (string)$sheet['name']
					);
					
				}
				$workbookRelationsXML = simplexml_load_string($this->getEntryData($workbookDir . '_rels/' . basename($rel['Target']) . '.rels'));
				foreach($workbookRelationsXML->Relationship as $wrel) {
					switch($wrel['Type']) {
						case self::SCHEMA_WORKSHEETRELATION:
							$sheets[(string)$wrel['Id']]['path'] = $workbookDir . (string)$wrel['Target'];
							break;
						case self::SCHEMA_SHAREDSTRINGS:
							$sharedStringsXML = simplexml_load_string($this->getEntryData($workbookDir . (string)$wrel['Target']));
							foreach($sharedStringsXML->si as $val) {
								if(isset($val->t)) {
									$this->sharedStrings[] = (string)$val->t;
								} elseif(isset($val->r)) {
									$this->sharedStrings[] = XLSXWorksheet::parseRichText($val);
								}
							}
							break;
					}
				}
			}
		}
		$this->sheetInfo = array();
		foreach($sheets as $rid=>$info) {
			$this->sheetInfo[$info['name']] = array(
				'sheetId' => $info['sheetId'],
				'rid' => $rid,
				'path' => $info['path']
			);
		}
	}

	// returns an array of sheet names, indexed by sheetId
	public function getSheetNames() {
		$res = array();
		foreach($this->sheetInfo as $sheetName=>$info) {
			$res[$info['sheetId']] = $sheetName;
		}
		return $res;
	}

	public function getSheetCount() {
		return count($this->sheetInfo);
	}

	// instantiates a sheet object (if needed) and returns an array of its data
	public function getSheetData($sheetNameOrId) {
		$sheet = $this->getSheet($sheetNameOrId);
		return $sheet->getData();
	}

	// instantiates a sheet object (if needed) and returns the sheet object
	public function getSheet($sheet) {
		if(is_numeric($sheet)) {
			$sheet = $this->getSheetNameById($sheet);
		} elseif(!is_string($sheet)) {
			throw new Exception("Sheet must be a string or a sheet Id");
		}
		if(!array_key_exists($sheet, $this->sheets)) {
			$this->sheets[$sheet] = new XLSXWorksheet($this->getSheetXML($sheet), $sheet, $this);

		}
		return $this->sheets[$sheet];
	}

	public function getSheetNameById($sheetId) {
		foreach($this->sheetInfo as $sheetName=>$sheetInfo) {
			if($sheetInfo['sheetId'] === $sheetId) {
				return $sheetName;
			}
		}
		throw new Exception("Sheet ID $sheetId does not exist in the Excel file");
	}

	protected function getSheetXML($name) {
		return simplexml_load_string($this->getEntryData($this->sheetInfo[$name]['path']));
	}

	// converts an Excel date field (a number) to a unix timestamp (granularity: seconds)
	public static function toUnixTimeStamp($excelDateTime) {
		if(!is_numeric($excelDateTime)) {
			return $excelDateTime;
		}
		$d = floor($excelDateTime); // seconds since 1900
		$t = $excelDateTime - $d;
		return ($d > 0) ? ( $d - 25569 ) * 86400 + $t * 86400 : $t * 86400;
	}

}

class XLSXWorksheet {

	protected $workbook;
	public $sheetName;
	protected $data;
	public $colCount;
	public $rowCount;
	protected $config;

	public function __construct($xml, $sheetName, XLSXReader $workbook) {
		$this->config = $workbook->config;
		$this->sheetName = $sheetName;
		$this->workbook = $workbook;
		$this->parse($xml);
	}

	// returns an array of the data from the sheet
	public function getData() {
		return $this->data;
	}

	protected function parse($xml) {
		$this->parseDimensions($xml->dimension);
		$this->parseData($xml->sheetData);
	}

	protected function parseDimensions($dimensions) {
		$range = (string) $dimensions['ref'];
		$cells = explode(':', $range);
		$maxValues = $this->getColumnIndex($cells[1]);
		$this->colCount = $maxValues[0] + 1;
		$this->rowCount = $maxValues[1] + 1;
	}

	protected function parseData($sheetData) {
		$rows = array();
		$curR = 0;
		$lastDataRow = -1;
		foreach ($sheetData->row as $row) {
			$rowNum = (int)$row['r'];
			if($rowNum != ($curR + 1)) {
				$missingRows = $rowNum - ($curR + 1);
				for($i=0; $i < $missingRows; $i++) {
					$rows[$curR] = array_pad(array(),$this->colCount,null);
					$curR++;
				}
			}
			$curC = 0;
			$rowData = array();
			foreach ($row->c as $c) {
				list($cellIndex,) = $this->getColumnIndex((string) $c['r']);
				if($cellIndex !== $curC) {
					$missingCols = $cellIndex - $curC;
					for($i=0;$i<$missingCols;$i++) {
						$rowData[$curC] = null;
						$curC++;
					}
				}
				$val = $this->parseCellValue($c);
				if(!is_null($val)) {
					$lastDataRow = $curR;
				}
				$rowData[$curC] = $val;
				$curC++;
			}
			$rows[$curR] = array_pad($rowData, $this->colCount, null);
			$curR++;
		}
		if($this->config['removeTrailingRows']) {
			$this->data = array_slice($rows, 0, $lastDataRow + 1);
			$this->rowCount = count($this->data);
		} else {
			$this->data = $rows;
		}
	}

	protected function getColumnIndex($cell = 'A1') {
		if (preg_match("/([A-Z]+)(\d+)/", $cell, $matches)) {
			
			$col = $matches[1];
			$row = $matches[2];
			$colLen = strlen($col);
			$index = 0;

			for ($i = $colLen-1; $i >= 0; $i--) {
				$index += (ord($col{$i}) - 64) * pow(26, $colLen-$i-1);
			}
			return array($index-1, $row-1);
		}
		throw new Exception("Invalid cell index");
	}
	
	protected function parseCellValue($cell) {
		// $cell['t'] is the cell type
		switch ((string)$cell["t"]) {
			case "s": // Value is a shared string
				if ((string)$cell->v != '') {
					$value = $this->workbook->sharedStrings[intval($cell->v)];
				} else {
					$value = '';
				}
				break;
			case "b": // Value is boolean
				$value = (string)$cell->v;
				if ($value == '0') {
					$value = false;
				} else if ($value == '1') {
					$value = true;
				} else {
					$value = (bool)$cell->v;
				}
				break;
			case "inlineStr": // Value is rich text inline
				$value = self::parseRichText($cell->is);
				break;
			case "e": // Value is an error message
				if ((string)$cell->v != '') {
					$value = (string)$cell->v;
				} else {
					$value = '';
				}
				break;
			default:
				if(!isset($cell->v)) {
					return null;
				}
				$value = (string)$cell->v;

				// Check for numeric values
				if (is_numeric($value)) {
					if ($value == (int)$value) $value = (int)$value;
					elseif ($value == (float)$value) $value = (float)$value;
					elseif ($value == (double)$value) $value = (double)$value;
				}
		}
		return $value;
	}

	// returns the text content from a rich text or inline string field
    public static function parseRichText($is = null) {
        $value = array();
        if (isset($is->t)) {
            $value[] = (string)$is->t;
        } else {
            foreach ($is->r as $run) {
                $value[] = (string)$run->t;
            }
        }
        return implode(' ', $value);
    }
}
