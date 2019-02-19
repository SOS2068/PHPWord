<?php
	require_once 'PHPWord.php';
	// New Word Document
	$PHPWord = new PHPWord();

	// New portrait section
	$section = $PHPWord->createSection();
	
	//addheader
	$header = $section->createHeader();

	//addfooter
	$footer = $section->createFooter();
	$footer->addPreserveText('{PAGE}');

	$table = $section->addTable();

	// connect DB
	$objConnect = mysqli_connect("localhost","root","","calmlversion2newform" ) or die ("error to connect DB");
	$Chapter1 = "select distinct
	chapter_des_2.chapter_code as Chapter,
	null as BlockCode,
	null as CategoryCode,
	chapter_des_2.code_kind as code_kind,
	chapter_des_2.preferred_label_id as preferred_label_id,
	chapter_des_2.preferred_label_description as preferred_label_description,
	null as BlockRange
	from chapter_des_2
	
	union all 
	
	select distinct
	parent as Chapter, /*from block_v2*/
	block_v2.block_code as BlockCode,
	null as CategoryCode,
	'block' as code_kind,
	block_v2.preferred_block_id as preferred_label_id ,/*from block_v2*/
	block_v2.preferred_block_description as preferred_label_description, /*from block_v2*/
	null as BlockRange
	from block_v2
	
	union all
	
	select distinct
	block_v2.parent as Chapter, /*from block_v2*/
	first_block_child.block_child as BlockCode,
	null as CategoryCode,
	'block' as code_kind,
	first_block_child.preferred_block_id as preferred_label_id ,/*from block_v2*/
	first_block_child.preferred_block_description as preferred_label_description, /*from block_v2*/
	null as BlockRange
	from first_block_child
	left join block_v2 on first_block_child.block_parent = block_v2.block_code
	
	union all
	
	select distinct
	block_v2.parent as Chapter, /*from block_v2*/
	second_block_child.block_child as BlockCode,
	null as CategoryCode,
	'block' as code_kind,
	second_block_child.preferred_block_id as preferred_label_id ,/*from block_v2*/
	second_block_child.preferred_block_description as preferred_label_description, /*from block_v2*/
	null as BlockRange
	from second_block_child
	left join first_block_child on second_block_child.block_parent = first_block_child.block_child
	left join block_v2 on first_block_child.block_parent = block_v2.block_code
	
	union all
	
	select distinct
	null as Chapter,
	category_v2.parent as BlockCode, 
	category_v2.category_code as CategoryCode,
	'category' as code_kind,
	category_v2.preferred_category_id as preferred_label_id, 
	category_v2.preferred_category_description as preferred_description,
	null as BlockRange
	from category_v2
	left join block_v2 on block_v2.block_code = category_v2.parent
	
	order by  preferred_label_id, BlockCode
	;";
	$chapter_1 = mysqli_query($objConnect,$Chapter1);

	// Add text elements
	$b = '';
	$d = '';
	$f = '';
	$h = '';
	$j = '';
	$incl = 'inclusion';
	$excl = 'exclusion';
	$codhin = 'coding-hint';
	$intro = 'introduction';
	$note = 'note';
	$text = 'text';
	$numpageI = 0;	
	
	while($chapterResult = mysqli_fetch_array($chapter_1,MYSQLI_ASSOC)){
			
			$a = $chapterResult["BlockCode"];
			$c = $chapterResult["code_kind"];
			$g = $chapterResult["CategoryCode"];
			
				if($chapterResult["code_kind"] == "chapter"){
					if($chapterResult["code_kind"] != $d){
						$section->addText($chapterResult["Chapter"]);
						$section->addText($chapterResult["preferred_label_description"]);
						$section->addText("(".($chapterResult["BlockRange"].")"));
					}
				}
				elseif($chapterResult["code_kind"] == "block"){
					if( $chapterResult["BlockCode"] != $b){
						if($chapterResult["code_kind"] == "block"){
							$section->addText($chapterResult["preferred_label_description"] ." "."(".$chapterResult["BlockCode"].")" );
						}
					}
				}
				elseif($chapterResult["code_kind"] == "category"){
					if( $chapterResult["BlockCode"] != $h){
						if($chapterResult["code_kind"] == "category"){
							$section->addText($chapterResult["CategoryCode"]." ".$chapterResult["preferred_label_description"]);
						}
					}
				}
				
			$b=$a;
			$d=$c;
			$h=$g;
			

	}	
	$section->addPageBreak();
	
	// Save File
	$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
	$objWriter->save('TabularlistVer2.docx');
?>