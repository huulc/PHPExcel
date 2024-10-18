<?php

function exportexcel() {
	global $wpdb;
	//Include PHPExcel
	// require_once (VISION_EXPORT_PLUGIN_DIR . "/lib/PHPExcel.php");
	include "PHPExcel.php";
	
	// Create new PHPExcel object
	$objPHPExcel = new PHPExcel();
	$objPHPExcel->setActiveSheetIndex(0);
	
	// Set document properties
	$headings = [
		'No',
		'post_title',

		'post_create',
		'_edit_last',
		'author_edit_last',

		'vision_faq_user_regis_non_regis',
		'number_post_views',
		'number_like_1',
		'number_like_2',
		'number_like_3',
		'ID',
		'post_content',
		'post_excerpt',
		'category_1',
		'category_2',
		'category_3',
		'category_4'

	];
	$i = 0;
	$keyArr = [];
	foreach ($headings as $index => $value) {
		if($value == 'No'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, 'No', \PHPExcel_Cell_DataType::TYPE_STRING);
		}elseif($value == 'post_title'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, 'タイトル', \PHPExcel_Cell_DataType::TYPE_STRING);// post_title
		}elseif($value == 'post_create'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, '日付', \PHPExcel_Cell_DataType::TYPE_STRING);// post_create
		}elseif($value == '_edit_last'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, '更新日時', \PHPExcel_Cell_DataType::TYPE_STRING);// _edit_last
		}elseif($value == 'author_edit_last'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, '更新者', \PHPExcel_Cell_DataType::TYPE_STRING);// author_edit_last
		}elseif($value == 'vision_faq_user_regis_non_regis'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, 'ご利用中/ご利用前', \PHPExcel_Cell_DataType::TYPE_STRING);// vision_faq_user_regis_non_regis
		}elseif($value == 'number_post_views'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, '閲覧数', \PHPExcel_Cell_DataType::TYPE_STRING);// number_post_views
		}elseif($value == 'number_like_1'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, '解決できた', \PHPExcel_Cell_DataType::TYPE_STRING);// number_like_1
		}elseif($value == 'number_like_2'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, '解決したが、納得できない', \PHPExcel_Cell_DataType::TYPE_STRING);// number_like_2
		}elseif($value == 'number_like_3'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, '役に立たない', \PHPExcel_Cell_DataType::TYPE_STRING);// number_like_3
		}elseif($value == 'ID'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, 'ID', \PHPExcel_Cell_DataType::TYPE_STRING);// ID
		}elseif($value == 'post_content'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, '回答', \PHPExcel_Cell_DataType::TYPE_STRING);// post_content
		}elseif($value == 'post_excerpt'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, 'キーワード', \PHPExcel_Cell_DataType::TYPE_STRING);// post_excerpt
		}elseif($value == 'category_1'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, '商材', \PHPExcel_Cell_DataType::TYPE_STRING);// category_1
		}elseif($value == 'category_2'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, 'カテゴリー1', \PHPExcel_Cell_DataType::TYPE_STRING);// category_2
		}elseif($value == 'category_3'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, 'カテゴリー2', \PHPExcel_Cell_DataType::TYPE_STRING);// category_3
		}elseif($value == 'category_4'){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($index, 1, 'カテゴリー3', \PHPExcel_Cell_DataType::TYPE_STRING);// category_4
		}
		$keyArr[] = $value;
		$i++;
	}
	
	$objPHPExcel->getActiveSheet()->getStyle('A1:M1')->getFont()->setBold(true);
	// $objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('A:M')->setAutoSize(true);


	$query = "SELECT * FROM {$wpdb->prefix}posts AS p WHERE p.post_type = 'vision-faq' AND p.post_status = 'publish' ORDER BY p.ID ASC";
	$posts   = $wpdb->get_results($query, 'ARRAY_A');
	
	if ( $posts ) {
		$arrCategory=[
			'0' => 'category0',
			'1' => 'category1',
			'2' => 'category2',
			'3' => 'category3',
		];
		foreach ( $posts as $key => $value ) {
			for ($j=0;$j<$i;$j++) {
					if($keyArr[$j] == 'No'){
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $key+1, \PHPExcel_Cell_DataType::TYPE_STRING);
					}else if($keyArr[$j] == 'post_title'){
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $value['post_title'], \PHPExcel_Cell_DataType::TYPE_STRING);
					}else if($keyArr[$j] == 'post_create'){
						$date_format = 'Y/m/d h:i:s';
						$date_create = get_the_date( $date_format, $value['ID'] );
						
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $date_create, \PHPExcel_Cell_DataType::TYPE_STRING);
					}else if($keyArr[$j] == '_edit_last'){
						$date_format = 'Y/m/d h:i:s';
						$date_edit = get_the_modified_date( $date_format, $value['ID'] );
						
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $date_edit, \PHPExcel_Cell_DataType::TYPE_STRING);
					}else if($keyArr[$j] == 'author_edit_last'){
						$last_id = get_post_meta( $value['ID'], '_edit_last', true );
						if ( $last_id ) {
							$last_user = get_userdata( $last_id );
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $last_user->display_name, \PHPExcel_Cell_DataType::TYPE_STRING);
						}else{
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, '', \PHPExcel_Cell_DataType::TYPE_STRING);
						}
					}else if($keyArr[$j] == 'vision_faq_user_regis_non_regis'){
						$vision_faq_user_regis_non_regis = get_post_meta( $value['ID'], 'vision_faq_user_regis_non_regis', true );
						if ( $vision_faq_user_regis_non_regis ) {
							if($vision_faq_user_regis_non_regis == 1){
								$string_regis = 'ご利用前';
							}else{
								$string_regis = 'ご利用中';
							}
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $string_regis, \PHPExcel_Cell_DataType::TYPE_STRING);
						}else{
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, '', \PHPExcel_Cell_DataType::TYPE_STRING);
						}
					}else if($keyArr[$j] == 'number_post_views'){
						$number_post_views = get_post_meta( $value['ID'], 'number_post_views', true );
						if ( $number_post_views ) {
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $number_post_views, \PHPExcel_Cell_DataType::TYPE_STRING);
						}else{
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, '', \PHPExcel_Cell_DataType::TYPE_STRING);
						}
					}else if($keyArr[$j] == 'number_like_1'){
						$number_like_1 = get_post_meta( $value['ID'], 'number_like_1', true );
						if ( $number_like_1 ) {
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $number_like_1, \PHPExcel_Cell_DataType::TYPE_STRING);
						}else{
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, '0', \PHPExcel_Cell_DataType::TYPE_STRING);
						}
					}else if($keyArr[$j] == 'number_like_2'){
						$number_like_2 = get_post_meta( $value['ID'], 'number_like_2', true );
						if ( $number_like_2 ) {
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $number_like_2, \PHPExcel_Cell_DataType::TYPE_STRING);
						}else{
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, '0', \PHPExcel_Cell_DataType::TYPE_STRING);
						}
					}else if($keyArr[$j] == 'number_like_3'){
						$number_like_3 = get_post_meta( $value['ID'], 'number_like_3', true );
						if ( $number_like_3 ) {
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $number_like_3, \PHPExcel_Cell_DataType::TYPE_STRING);
						}else{
							$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, '0', \PHPExcel_Cell_DataType::TYPE_STRING);
						}
					}else if($keyArr[$j] == 'ID'){
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $value['ID'], \PHPExcel_Cell_DataType::TYPE_STRING);
					}else if($keyArr[$j] == 'post_content'){
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $value['post_content'], \PHPExcel_Cell_DataType::TYPE_STRING);
						// $itinerary = nl2br( $value['post_content'] );
						// $itinerary = preg_replace( "#<br />(\r\n|\n|\r)#", '<br />', $itinerary );
						// $itinerary = str_replace( '"' , "'" , $itinerary );
						// $itinerary = strip_tags($itinerary);
						// $csv_output .= '"'.$itinerary .'"	';

						// $itinerary = sanitize_textarea_field( $value['post_content'] );
						// $itinerary = nl2br( $itinerary );
						// $itinerary = preg_replace( "#<br />(\r\n|\n|\r)#", '<br />', $itinerary );
						// $csv_output .= '"'.$itinerary .'"	';
					}else if($keyArr[$j] == 'post_excerpt'){
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $value['post_excerpt'], \PHPExcel_Cell_DataType::TYPE_STRING);
					}else if($keyArr[$j] == 'category_1'){
						if(!empty($arrCategory['0'])){
							$catName = $arrCategory['0'];
						}else{
							$catName = '';
						}
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $catName, \PHPExcel_Cell_DataType::TYPE_STRING);
					}else if($keyArr[$j] == 'category_2'){
						if(!empty($arrCategory['1'])){
							$catName = $arrCategory['1'];
						}else{
							$catName = '';
						}
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $catName, \PHPExcel_Cell_DataType::TYPE_STRING);
					}else if($keyArr[$j] == 'category_3'){
						if(!empty($arrCategory['2'])){
							$catName = $arrCategory['2'];
						}else{
							$catName = '';
						}
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $catName, \PHPExcel_Cell_DataType::TYPE_STRING);
					}else if($keyArr[$j] == 'category_4'){
						if(!empty($arrCategory['3'])){
							$catName = $arrCategory['3'];
						}else{
							$catName = '';
						}
						$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($j, $key+2, $catName, \PHPExcel_Cell_DataType::TYPE_STRING);
					}
				
			}
		}
	}

	// Rename worksheet
	//$objPHPExcel->getActiveSheet()->setTitle('Simple');
	
	// Set active sheet index to the first sheet, so Excel opens this as the first sheet
	// $objPHPExcel->setActiveSheetIndex(0);
	set_time_limit(0);
	ini_set('memory_limit', '-1');
	// Redirect output to a client’s web browser
	ob_clean();
	ob_start();
	$file = 'export_qa';
	$filename = $file."_".date("M-d-Y");
	switch ( $_GET['format'] ) {
		case 'csv':
			// Redirect output to a client’s web browser (CSV)
			header("Content-type: text/csv");
			header("Cache-Control: no-store, no-cache");
			header('Content-Disposition: attachment; filename="'.$filename.'.csv"');
			$objWriter = new PHPExcel_Writer_CSV($objPHPExcel);
			$objWriter->setDelimiter(',');
			$objWriter->setEnclosure('"');
			$objWriter->setLineEnding("\r\n");
			//$objWriter->setUseBOM(true);
			$objWriter->setSheetIndex(0);
			$objWriter->save('php://output');
			break;
		case 'xls':				
			// Redirect output to a client’s web browser (Excel5)
			header('Content-Type: application/vnd.ms-excel');
			header('Content-Disposition: attachment;filename="'.$filename.'.xls"');
			header('Cache-Control: max-age=0');
			
			$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
			$objWriter->save('php://output');
			break;
		case 'xlsx':
			// Redirect output to a client’s web browser (Excel2007)
			header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
			header('Content-Disposition: attachment;filename="'.$filename.'.xlsx"');
			header('Cache-Control: max-age=0');
			$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
			$objWriter->save('php://output');
			break;
	}
	exit;

}