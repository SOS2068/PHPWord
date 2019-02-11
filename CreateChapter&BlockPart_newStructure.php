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
	null  as BlockStatment,
	null as BlockCode, 
	null as CategoryCode,
	null as SubCategory,
	min(category_v2.category_code) as minCat,
	max(category_v2.category_code) as maxCat,
	code_kind,
	preferred_label_id,
	preferred_label_description as preferred_description, 
	null as label_id,
	null as description_type,
	null as label_description
	from chapter_des_2
	left join block_v2 on block_v2.parent = chapter_des_2.chapter_code
	left join category_v2 on category_v2.parent = block_v2.block_code
	where chapter_des_2.chapter_code = 'I'
	union all 
	select distinct 
	chapter_code as Chapter,
	null  as BlockStatment,
	null as BlockCode, 
	null as CategoryCode,
	null as SubCategory,
	null as minCat,
	null as maxCat,
	code_kind, 
	preferred_label_id, 
	preferred_label_description as preferred_description, 
	label_id, 
	description_type, 
	label_description  
	from chapter_des_2
	where chapter_des_2.chapter_code = 'I'
	union all
	select distinct
	parent as Chapter, /*from block_v2*/
	block_v2.block_code as BlockStatment,
	null as BlockCode,
	null as CategoryCode,
	null as SubCategory,
	null as minCat,
	null as maxCat,
	'blockStatement' as code_kind,
	'D0000002' as preferred_label_id ,/*from block_v2*/
	block_v2.preferred_block_description as preferred_label_description, /*from block_v2*/
	null as label_id,
	null as description_type,
	null as label_description
	from block_v2
	left join block_des_v2 on block_des_v2.parent_id = block_v2.preferred_block_id
	where block_v2.parent = 'I'
	union all
	select distinct 
	parent as Chapter, /*from block_v2*/
	null  as BlockStatment,
	block_v2.block_code as BlockCode,
	null as CategoryCode, /*from block_v2*/
	null as SubCategory,
	null as minCat,
	null as maxCat,
	'block' as code_kind,
	preferred_block_id as preferred_label_id ,/*from block_v2*/
	block_v2.preferred_block_description as preferred_label_description, /*from block_v2*/
	label_id,
	block_des_v2.label_type as description_type,
	label_description
	from block_v2
	left join block_des_v2 on block_des_v2.parent_id = block_v2.preferred_block_id
	where block_v2.parent = 'I'
	union all
	select distinct
	null as Chapter,
	null  as BlockStatment,
	category_v2.parent as BlockCode, /*from category-v2*/
	category_v2.category_code as CategoryCode,
	null as SubCategory,
	null as minCat,
	null as maxCat,
	'category' as code_kind,
	category_v2.preferred_category_id as preferred_label_id, /*from category-v2*/
	category_v2.preferred_category_description as preferred_description, /*from category-v2*/
	category_des_v2.category_id as label_id,
	category_des_v2.label_type as description_type,
	category_des_v2.category_description as label_description
	from category_v2
	left join category_des_v2 on category_des_v2.preferred_category_id = category_v2.preferred_category_id
	left join block_v2 on block_v2.block_code = category_v2.parent
	where block_v2.parent = 'I'
	union all
	select distinct
	null as Chapter,
	null  as BlockStatment,
	block_v2.block_code as BlockCode,
	sub_category_v2.parent as CategoryCode,
	sub_category_v2.sub_category_code as SubCategory,
	null as minCat,
	null as maxCat,
	'subcategory' as code_kind,
	sub_category_v2.preferred_sub_category_id as preferred_label_id,
	sub_category_v2.preferred_sub_category_description as preferred_description,
	sub_category_des_v2.sub_category_id as label_id,
	sub_category_des_v2.label_type as description_type,
	sub_category_des_v2.sub_category_description as label_description
	from sub_category_v2
	left join sub_category_des_v2 on sub_category_des_v2.sub_category = sub_category_v2.preferred_sub_category_id
	left join category_v2 on category_v2.category_code = sub_category_v2.parent
	left join block_v2 on block_v2.block_code = category_v2.parent 
	where block_v2.parent = 'I'
	order by  BlockCode, preferred_label_id, label_id
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
			$g = $chapterResult["CategoryCode"];
			$i = $chapterResult["SubCategory"];
			
				if($chapterResult["code_kind"] == "chapter"){
					if($chapterResult["code_kind"] != $d){
						$section->addText("CHAPTER ".$chapterResult["Chapter"]);
						$section->addText($chapterResult["preferred_description"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
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
			$h=$g;
			$j=$i;

	}	
	$section->addPageBreak();
	
	// Save File
	$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
	$objWriter->save('Chapter&Block.docx');
?>