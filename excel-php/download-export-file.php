<?php

/** Error reporting */
error_reporting(E_ALL);
//ini_set('display_errors', TRUE);
//ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
	die('This example should only be run from a Web Browser');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/lib/PHPExcel/Classes/PHPExcel.php';

require_once("../../../wp-load.php");

if(!function_exists('return_order_product_id')) {
	// Query orders with product id
	function return_order_product_id($product_id) {
		global $wpdb;
		$tabelaOrderItemMeta = $wpdb->prefix . 'woocommerce_order_itemmeta';
		$tabelaOrderItems = $wpdb->prefix . 'woocommerce_order_items';

		$resultadoSelect = $wpdb->get_results(
			$wpdb->prepare(
				"SELECT b.order_id
				   FROM {$tabelaOrderItemMeta} a, {$tabelaOrderItems} b
				  WHERE a.meta_key = '_product_id'
					AND a.meta_value = %s
					AND a.order_item_id = b.order_item_id
					ORDER BY b.order_id DESC",
				$product_id
			)
		);

		if($resultadoSelect)
		{
			$result = array();

			foreach($resultadoSelect as $item)
				array_push($result, $item->order_id);

			if($result)
			{
				return $result;
			}
		}
	}
}

if( !isset($_GET['event']) ) {
	exit;
}

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Set document properties
$objPHPExcel->getProperties()
    ->setCreator("Five by Five")
	->setLastModifiedBy("Five by Five")
	->setTitle("GSCC Order Export - ". get_the_title($_GET['event']))
	->setSubject("GSCC Order Export - ". get_the_title($_GET['event']))
	->setDescription("GSCC Order Export - ". get_the_title($_GET['event']))
	->setKeywords("GSCC Order Export - ". get_the_title($_GET['event']))
	->setCategory("GSCC Order Export");
	
// Headers
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue('A1', 'Order ID')
	->setCellValue('B1', 'Order Date')
	->setCellValue('C1', 'First Name')
	->setCellValue('D1', 'Last Name')
	->setCellValue('E1', 'Requirements')
	->setCellValue('F1', 'Company')
	->setCellValue('G1', 'Email')
	->setCellValue('H1', 'Member');
	
$objPHPExcel->getActiveSheet()->getStyle('A1:H1')->getFont()->setBold(true); 


$event = $_GET['event'];

// get order ids that have the product id
$order_ids = return_order_product_id($event);

if( $order_ids ) {
	$query_orders = get_posts(array(
		'post_type' => 'shop_order',
		'post_status' => array('wc-processing', 'wc-completed'),
		'showposts' => -1,
		'post__in' => $order_ids,
	));
	
	$entries = array();
	
	$total_persons = 0;
	
	foreach( $query_orders as $qo ) {

		$meta = get_post_meta($qo->ID);
		$persons = array();
		
		$persons[] = array(
			'first_name' =>  $meta['_billing_first_name'][0],
			'last_name' =>  $meta['_billing_last_name'][0],
			'email' =>  $meta['_billing_email'][0],
			'company' =>  $meta['_billing_company'][0],
			'requirements' =>  $meta['_billing_dietary_requirements'][0],
			'member' =>  $meta['_billing_member_gscc'][0],
		);
		$total_persons++;
		
		// Additional Persons
		$order_obj = new WC_Order( $qo->ID );
		$order_items = $order_obj->get_items();
		$event_count = 0;
		
		foreach ($order_items as $order_item) {
			$event_count++;
			$quantity = $order_item['quantity'];
			$product_id = $order_item['product_id'];
			if( $event == $product_id ){
				for ($i=1; $i <= $quantity; $i++) {
					if( $meta['Event '.$event_count.' Person '. $i .' First Name'][0] ) {
						$total_persons++;
						$persons[] = array(
							'first_name' =>  $meta['Event '.$event_count.' Person '. $i .' First Name'][0],
							'last_name' =>  $meta['Event '.$event_count.' Person '. $i .' Last Name'][0],
							'requirements' =>  $meta['Event '.$event_count.' Person '. $i .' Dietary Requirements'][0],
						);
					}
				}
			}
		}

		$entries[] = array(
			'order_id' => $qo->ID,
			'order_date' => $qo->post_date,
			'persons' => $persons,
		);
		
	}
	
	$row = 1;
	foreach( $entries as $e ) {
		
		foreach( $e['persons'] as $p ) { $row++;
		
			$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue('A'. $row, $e['order_id'])
			->setCellValue('B'. $row, $e['order_date'])
			->setCellValue('C'. $row, $p['first_name'])
			->setCellValue('D'. $row, $p['last_name'])
			->setCellValue('E'. $row, $p['requirements'])
			->setCellValue('F'. $row, $p['company'])
			->setCellValue('G'. $row, $p['email'])
			->setCellValue('H'. $row, $p['member']);
			
		}
		
	}
	
	$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
	$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
	$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
	$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
	$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
	$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
	$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
	
}

// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('Simple');

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);

// Redirect output to a client’s web browser (Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="gscc-orders-'. basename(get_permalink($_GET['event'])) .'.xlsx"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');

exit;