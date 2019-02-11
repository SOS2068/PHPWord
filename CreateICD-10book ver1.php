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
	$strSQL1 = "SELECT * FROM rubric1";
	$strSQL2 = "SELECT class2.parent, rubric2.code, rubric2.label  from class2 right join rubric2 on class2.code=rubric2.code";
	$strSQL3 = "SELECT class3.parent, rubric3.code, rubric3.label  from class3 right join rubric3 on class3.code=rubric3.code";
	$strSQL4 = "SELECT class4.parent, rubric4.code, rubric4.label  from class4 right join rubric4 on class4.code=rubric4.code";
	$chapter = mysqli_query($objConnect,$strSQL1);
	$block = mysqli_query($objConnect,$strSQL2);
	$category = mysqli_query($objConnect,$strSQL3);
	$codelv1 = mysqli_query($objConnect,$strSQL4);

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
				$section->addText($chapterResult["code"]);
			}

			if($chapterResult["kind" ]== $incl){
				$section->addListItem($chapterResult["label"],1);
			}elseif($chapterResult["kind"] == $codhin){
				$section->addListItem($chapterResult["label"],1);
			}elseif($chapterResult["kind"] == $excl){
				$section->addListItem($chapterResult["label"],1);
			}elseif($chapterResult["kind"] == $intro){
				$section->addListItem($chapterResult["label"],1);
			}elseif($chapterResult["kind"] == $note){
				$section->addListItem($chapterResult["label"],1);
			}elseif($chapterResult["kind" ]== $text){
				$section->addListItem($chapterResult["label"],1);
			}else{
				$section->addText($chapterResult["label"]);
			}


			$b=$a;
	}
	// Save File
	$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
	$objWriter->save('test.docx');
?>