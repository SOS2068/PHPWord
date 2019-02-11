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
	$objConnect = mysqli_connect("localhost","root","","calmlversion2" ) or die ("error to connect DB");
	$Chapter1 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='I' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='I' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'I' and rubric2new1.kind = 'preferred'
	;";
	$Chapter2 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='II' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='II' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'II' and rubric2new1.kind = 'preferred'
	;";
	$Chapter3 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='III' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='III' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'III' and rubric2new1.kind = 'preferred'
	;";
	$Chapter4 ="select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='IV' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='IV' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'IV' and rubric2new1.kind = 'preferred'
	;";
	$Chapter5 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='V' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='V' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'v' and rubric2new1.kind = 'preferred'
	;";
	$Chapter6 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='VI' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='VI' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'VI' and rubric2new1.kind = 'preferred'
	;";
	$Chapter7 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='VII' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='VII' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'VII' and rubric2new1.kind = 'preferred'
	;";
	$Chapter8 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='VIII' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='VIII' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'VIII' and rubric2new1.kind = 'preferred'
	;";
	$Chapter9 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='IX' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='IX' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'IX' and rubric2new1.kind = 'preferred'
	;";
	$Chapter10 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='X' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='X' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'X' and rubric2new1.kind = 'preferred'
	;";
	$Chapter11 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XI' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XI' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XI' and rubric2new1.kind = 'preferred'
	;";
	$Chapter12 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XII' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XII' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XII' and rubric2new1.kind = 'preferred'
	;";
	$Chapter13 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XIII' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XIII' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XIII' and rubric2new1.kind = 'preferred'
	;";
	$Chapter14 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XIV' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XIV' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XIV' and rubric2new1.kind = 'preferred'
	;";
	$Chapter15 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XV' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XV' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XV' and rubric2new1.kind = 'preferred'
	;";
	$Chapter16 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XVI' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XVI' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XVI' and rubric2new1.kind = 'preferred'
	;";
	$Chapter17 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XVII' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XVII' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XVII' and rubric2new1.kind = 'preferred'
	;";
	$Chapter18 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XVIII' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XVIII' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XVIII' and rubric2new1.kind = 'preferred'
	;";
	$Chapter19 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XIX' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XIX' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XIX' and rubric2new1.kind = 'preferred'
	;";
	$Chapter20 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XX' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XX' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XX' and rubric2new1.kind = 'preferred'
	;";
	$Chapter21 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XXI' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XXI' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XXI' and rubric2new1.kind = 'preferred'
	;";
	$Chapter22 = "select distinct class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, min(class3.code) as minCat,  max(class3.code) as maxCat, class2.code as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	right join class2 on class2.parent = rubric1new1.code
	right join class3 on class3.parent = class2.code
	where rubric1new1.code ='XXII' and rubric1new1.kind = 'preferred'
	union all
	select distinct  class1.parent, class1.kind as classkide ,rubric1new1.code, rubric1new1.kind, rubric1new1.label, null as minCat, null as maxCat, null as borrow from class1
	right join rubric1new1 on class1.code = rubric1new1.code
	where rubric1new1.code ='XXII' and not rubric1new1.kind = 'preferred'
	union all
	select distinct class2.parent, class2.kind as classkide, rubric2new1.code,rubric2new1.kind, rubric2new1.label, null as minCat, null as maxCat, null as borrow from class2
	right join rubric2new1 on class2.code = rubric2new1.code
	right join class3 on class3.parent = rubric2new1.code
	where class2.parent = 'XXII' and rubric2new1.kind = 'preferred'
	;";
	
	$chapter_1 = mysqli_query($objConnect,$Chapter1);	$chapter_12 = mysqli_query($objConnect,$Chapter12);
	$chapter_2 = mysqli_query($objConnect,$Chapter2);	$chapter_13 = mysqli_query($objConnect,$Chapter13);
	$chapter_3 = mysqli_query($objConnect,$Chapter3);	$chapter_14 = mysqli_query($objConnect,$Chapter14);
	$chapter_4 = mysqli_query($objConnect,$Chapter4);	$chapter_15 = mysqli_query($objConnect,$Chapter15);
	$chapter_5 = mysqli_query($objConnect,$Chapter5);	$chapter_16 = mysqli_query($objConnect,$Chapter16);
	$chapter_6 = mysqli_query($objConnect,$Chapter6);	$chapter_17 = mysqli_query($objConnect,$Chapter17);
	$chapter_7 = mysqli_query($objConnect,$Chapter7);	$chapter_18 = mysqli_query($objConnect,$Chapter18);
	$chapter_8 = mysqli_query($objConnect,$Chapter8);	$chapter_19 = mysqli_query($objConnect,$Chapter19);
	$chapter_9 = mysqli_query($objConnect,$Chapter9);	$chapter_20 = mysqli_query($objConnect,$Chapter20);
	$chapter_10 = mysqli_query($objConnect,$Chapter10);	$chapter_21 = mysqli_query($objConnect,$Chapter21);
	$chapter_11 = mysqli_query($objConnect,$Chapter11);	$chapter_22 = mysqli_query($objConnect,$Chapter22);

	// Add text elements
	$b = '';
	$incl = 'inclusion';
	$excl = 'exclusion';
	$codhin = 'coding-hint';
	$intro = 'introduction';
	$note = 'note';
	$text = 'text';
	$numpageI = 0;	
	
	while($chapterResult = mysqli_fetch_array($chapter_1,MYSQLI_ASSOC)){
			
			$a = $chapterResult["code"];
				if( $chapterResult["code"] != $b){
					if($chapterResult["classkide"] == "chapter"){
						$section->addText("CHAPTER ".$chapterResult["code"]);
						$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
					}
					else{
						$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
					}
				}
			
				if($chapterResult["kind" ]== $incl){
					$section->addText("	".$chapterResult["label"]);
				}elseif($chapterResult["kind"] == $codhin){
					$section->addText("	".$chapterResult["label"],1);
				}elseif($chapterResult["kind"] == $excl){
					$section->addText("	".$chapterResult["label"],1);
				}elseif($chapterResult["kind"] == $intro){
					$section->addText("	".$chapterResult["label"],1);
				}elseif($chapterResult["kind"] == $note){
					$section->addText("	".$chapterResult["label"],1);
				}elseif($chapterResult["kind" ]== $text){
					$section->addText("	".$chapterResult["label"],1);
				}
			$b=$a;	
	}	
	$section->addPageBreak();

	while($chapterResult = mysqli_fetch_array($chapter_2,MYSQLI_ASSOC)){
			
		$a = $chapterResult["code"];
			if( $chapterResult["code"] != $b){
				if($chapterResult["classkide"] == "chapter"){
					$section->addText("CHAPTER ".$chapterResult["code"]);
					$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
				}
				else{
					$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
				}
			}
		
			if($chapterResult["kind" ]== $incl){
				$section->addText("	".$chapterResult["label"]);
			}elseif($chapterResult["kind"] == $codhin){
				$section->addText("	".$chapterResult["label"],1);
			}elseif($chapterResult["kind"] == $excl){
				$section->addText("	".$chapterResult["label"],1);
			}elseif($chapterResult["kind"] == $intro){
				$section->addText("	".$chapterResult["label"],1);
			}elseif($chapterResult["kind"] == $note){
				$section->addText("	".$chapterResult["label"],1);
			}elseif($chapterResult["kind" ]== $text){
				$section->addText("	".$chapterResult["label"],1);
			}
		$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_3,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("		".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_4,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_5,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_6,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_7,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_8,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_9,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_10,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_11,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_12,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_13,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_14,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_15,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_16,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_17,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_18,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_19,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_20,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_21,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."	".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();

while($chapterResult = mysqli_fetch_array($chapter_22,MYSQLI_ASSOC)){
			
	$a = $chapterResult["code"];
		if( $chapterResult["code"] != $b){
			if($chapterResult["classkide"] == "chapter"){
				$section->addText("CHAPTER ".$chapterResult["code"]);
				$section->addText($chapterResult["label"]." "."(".$chapterResult["minCat"]."-".$chapterResult["maxCat"].")");
			}
			else{
				$section->addText($chapterResult["code"]."		".$chapterResult["label"]);
			}
		}
	
		if($chapterResult["kind" ]== $incl){
			$section->addText("	".$chapterResult["label"]);
		}elseif($chapterResult["kind"] == $codhin){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $excl){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $intro){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind"] == $note){
			$section->addText("	".$chapterResult["label"],1);
		}elseif($chapterResult["kind" ]== $text){
			$section->addText("	".$chapterResult["label"],1);
		}
	$b=$a;	
}	
$section->addPageBreak();
	
	// Save File
	$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
	$objWriter->save('Chapter.docx');
?>