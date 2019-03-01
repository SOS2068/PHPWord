<?php
	require_once 'PHPWord.php';

	// New Word Document
	$PHPWord = new PHPWord();

	// New portrait section
	$section = $PHPWord->createSection();

	$table = $section->addTable();

	
	$objConnect = mysqli_connect("localhost","root","","calmlversion3newform" ) or die ("error to connect DB");
	$Chapter1 = "select distinct sub_category_v3.sub_category_code as code ,sub_category_v3.parent
	from sub_category_v3;";
	
	$chapter_1 = mysqli_query($objConnect,$Chapter1);
	

	// Add text elements
	$b = '';
	$count = 0;
	
	$table->addRow(900);
	$table->addCell(2000)->addText('Code');
	$table->addCell(2000)->addText('Number');
	while($chapterResult = mysqli_fetch_array($chapter_1,MYSQLI_ASSOC)){
			
			$a = $chapterResult["parent"];
				if( $chapterResult["parent"] != $b){
					$table->addRow();
					$table->addCell(2000)->addText($b);
					$table->addCell(2000)->addText($count);
					$count = 0;
				}
				$count = $count+1;
				$b=$a;

	}	

	
	// Save File
	$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
	$objWriter->save('count.docx');
?>