<?php
	require_once 'PHPWord.php';

	// New Word Document
	$PHPWord = new PHPWord();

	// New portrait section
	$section = $PHPWord->createSection();

	$PHPWord->addFontStyle('rStyle', array('bold'=>true, 'italic'=>true, 'size'=>16));
	$PHPWord->addParagraphStyle('pStyle', array('align'=>'center', 'spaceAfter'=>100));
	$section->addText('Customer Report', 'rStyle', 'pStyle');

	// Define table style arrays
	$styleTable = array('borderSize'=>6, 'borderColor'=>'006699', 'cellMargin'=>80);
	$styleFirstRow = array('borderBottomSize'=>18, 'borderBottomColor'=>'0000FF', 'bgColor'=>'66BBFF');

	// Define cell style arrays
	$styleCell = array('valign'=>'center');
	$styleCellBTLR = array('valign'=>'center', 'textDirection'=>PHPWord_Style_Cell::TEXT_DIR_BTLR);

	// Define font style for first row
	$fontStyle = array('bold'=>true, 'align'=>'center');

	// Add table style
	$PHPWord->addTableStyle('myOwnTableStyle', $styleTable, $styleFirstRow);

	// Add table
	$table = $section->addTable('myOwnTableStyle');

	// Add row
	$table->addRow(200);

	// Add cells
	$table->addCell(1500, $styleCell)->addText('CustomerID', $fontStyle);
	$table->addCell(1500, $styleCell)->addText('Name', $fontStyle);
	$table->addCell(1500, $styleCell)->addText('Email', $fontStyle);
	$table->addCell(1500, $styleCell)->addText('CountryCode', $fontStyle);
	$table->addCell(1500, $styleCell)->addText('Budget', $fontStyle);
	$table->addCell(1500, $styleCell)->addText('Used', $fontStyle);

	// Write data from MySQL result
	$objConnect = mysql_connect("localhost","root","root") or die("Error Connect to Database");
	$objDB = mysql_select_db("mydatabase");
	$strSQL = "SELECT * FROM customer";
	$objQuery = mysql_query($strSQL);
	while($objResult = mysql_fetch_array($objQuery))
	{
			// Add more rows / cells
			$table->addRow();
			$table->addCell(1500)->addText($objResult["CustomerID"]);
			$table->addCell(1500)->addText($objResult["Name"]);
			$table->addCell(1500)->addText($objResult["Email"]);
			$table->addCell(1500)->addText($objResult["CountryCode"]);
			$table->addCell(1500)->addText($objResult["Budget"]);
			$table->addCell(1500)->addText($objResult["Used"]);
		}

	// Save File
	$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
	$objWriter->save('CreateWord2.docx');
?>