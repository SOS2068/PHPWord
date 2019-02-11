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
	$objConnect = mysqli_connect("localhost","root","","calmlversion3newform" ) or die ("error to connect DB");
	$Chapter1 = "select distinct
	chapter_v3_blockrange.chapter_name as Chapter,
	null as BlockStatment,
	null as BlockCode,
	'chapter' as code_kind,
	chapter_v3_blockrange.preferred_label_id,
	chapter_v3_blockrange.preferred_label_description as preferred_description,
	chapter_v3_blockrange.block_range as BlockRange,
	null as label_id,
	null as description_type,
	null as label_description
	from chapter_v3_blockrange
	
	union all 
	
	select distinct 
	chapter_code as Chapter,
	null as BlockStatment,
	null as BlockCode, 
	code_kind, 
	preferred_label_id, 
	preferred_label_description as preferred_description, 
	null as BlockRange,
	label_id, 
	description_type, label_description  
	from chapter_des_v3
	
	union all
	
	select distinct
	parent as Chapter, /*from block_v2*/
	block_v3.block_code as BlockStatment,
	null as BlockCode,
	'blockStatement' as code_kind,
	block_v3.preferred_block_id as preferred_label_id ,/*from block_v2*/
	block_v3.preferred_block_description as preferred_label_description, /*from block_v2*/
	null as BlockRange,
	null as label_id,
	null as description_type,
	null as label_description
	from block_v3
	left join block_des_v3 on block_des_v3.preferred_block_id = block_v3.preferred_block_id
	
	union all
	
	select distinct
	block_v3.parent as Chapter, /*from block_v2*/
	block_first_child.block_code as BlockStatment,
	null as BlockCode,
	'blockStatement' as code_kind,
	block_first_child.preferred_block_id as preferred_label_id ,/*from block_v2*/
	block_first_child.preferred_block_description as preferred_label_description, /*from block_v2*/
	null as BlockRange,
	null as label_id,
	null as description_type,
	null as label_description
	from block_first_child
	left join block_v3 on block_first_child.preferred_block_parent = block_v3.block_code
	
	union all
	
	select distinct
	block_v3.parent as Chapter, /*from block_v2*/
	block_second_child.block_code as BlockStatment,
	null as BlockCode,
	'blockStatement' as code_kind,
	block_second_child.preferred_block_id as preferred_label_id ,/*from block_v2*/
	block_second_child.preferred_block_description as preferred_label_description, /*from block_v2*/
	null as BlockRange,
	null as label_id,
	null as description_type,
	null as label_description
	from block_second_child
	left join block_first_child on block_second_child.preferred_block_parent = block_first_child.block_code
	left join block_v3 on block_first_child.preferred_block_parent = block_v3.block_code
	
	
	order by  preferred_label_id, BlockStatment
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
			$e = $chapterResult["BlockStatment"];
		
			
				if($chapterResult["code_kind"] == "chapter"){
					if($chapterResult["code_kind"] != $d){
						$section->addText($chapterResult["Chapter"]);
						$section->addText($chapterResult["preferred_description"]);
						$section->addText("(".($chapterResult["BlockRange"].")"));
					}
				}
				elseif($chapterResult["code_kind"] == "blockStatement"){
					if($chapterResult["BlockStatment"] != $f){
						if($chapterResult["code_kind"] == "blockStatement"){
							$section->addText($chapterResult["BlockStatment"]."	".$chapterResult["preferred_description"]);
						}
					}
				}
				elseif($chapterResult["code_kind"] == "block"){
					if( $chapterResult["BlockCode"] != $b){
						if($chapterResult["code_kind"] == "block"){
							$section->addText($chapterResult["preferred_description"]);
							$section->addText($chapterResult["BlockCode"]);
						}
					}
				}
				elseif($chapterResult["code_kind"] == "category"){
					if( $chapterResult["BlockCode"] != $h){
						if($chapterResult["code_kind"] == "category"){
							$section->addText($chapterResult["CategoryCode"]." ".$chapterResult["preferred_description"]);
						}
					}
				}
				elseif($chapterResult["code_kind"] == "subcategory"){
					if( $chapterResult["SubCategory"] != $j){
						if($chapterResult["code_kind"] == "subcategory"){
							$section->addText($chapterResult["SubCategory"]." ".$chapterResult["preferred_description"]);
						}
					}
				}
				

				if($chapterResult["description_type" ]== $incl){
					$section->addText("	".$chapterResult["label_description"]);
				}elseif($chapterResult["description_type"] == $codhin){
					$section->addText("	".$chapterResult["label_description"],1);
				}elseif($chapterResult["description_type"] == $excl){
					$section->addText("	".$chapterResult["label_description"],1);
				}elseif($chapterResult["description_type"] == $intro){
					$section->addText("	".$chapterResult["label_description"],1);
				}elseif($chapterResult["description_type"] == $note){
					$section->addText("	".$chapterResult["label_description"],1);
				}elseif($chapterResult["description_type" ]== $text){
					$section->addText("	".$chapterResult["label_description"],1);
				}
			$b=$a;
			$d=$c;
			$f=$e;
			

	}	
	$section->addPageBreak();
	
	// Save File
	$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
	$objWriter->save('Chapter&BlockVer3.docx');
?>