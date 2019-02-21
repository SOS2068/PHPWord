<?php
	require_once 'PHPWord.php';

	// New Word Document
	$PHPWord = new PHPWord();
	// New portrait section
	$section = $PHPWord->createSection();
	
	//addheader
	$header = $section->createHeader();
	$header->addPreserveText('INTERNATIONAL CLASSIFICATION OF DISEASE', array('align'=>'center'));

	//addfooter
	$footer = $section->createFooter();
	$footer->addPreserveText('{NUMPAGES}', array('align'=>'center'));

	$table = $section->addTable();
	// connect DB
	$objConnect = mysqli_connect("localhost","root","","calml" ) or die ("error to connect DB");
	$strSQL1 = "select distinct  rubric1.code, rubric1.kind, rubric1.label, rubric1.id  from class1
	right join rubric1 on class1.code = rubric1.code
	where rubric1.kind = 'preferred'
	union all
	select distinct  rubric2.code,rubric2.kind, rubric2.label, rubric2.id from class2
	right join rubric2 on class2.code = rubric2.code
	where rubric2.kind = 'preferred'
	union all
	select distinct rubric3.code, rubric3.kind, rubric3.label, rubric3.id from class3
	right join rubric3 on class3.code = rubric3.code
	where rubric3.kind = 'preferred' 
	union all 
	select distinct rubric4.code, rubric4.kind, rubric4.label, rubric4.id from class4
	right join rubric4 on class4.parent = rubric4.code
	where rubric4.kind = 'preferred' and rubric4.code not like '%.%'
	union all
	select distinct rubric5.code, rubric5.kind, rubric5.label, rubric5.id from class5
	right join rubric5 on class5.parent = rubric5.code
	where rubric5.kind = 'preferred' and rubric5.code not like '%.%'
	
	order by id
	;
	";
	$chapter = mysqli_query($objConnect,$strSQL1);
	// Add text elements
	$b = '';
	$incl = 'inclusion';
	$excl = 'exclusion';
	$codhin = 'coding-hint';
	$intro = 'introduction';
	$note = 'note';
	$text = 'text';
	while($chapterResult = mysqli_fetch_array($chapter,MYSQLI_ASSOC)){
			$a = $chapterResult["code"];
			
			if( $chapterResult["code"] != $b){
				$section->addText($chapterResult["code"]."	". $chapterResult["label"]);
			}


			$b=$a;
	}
	// Save File
	$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
	$objWriter->save('TabularlistVer1.docx');
?>