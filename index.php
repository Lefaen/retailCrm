<?

require_once '/vendor/autoload.php';


$client = new \RetailCrm\ApiClient(
    'https://u5904sbar-mn1-justhost.retailcrm.ru',
    'hcYFnxKesEvibWotC9yAkJstLxMw0FdO',
    \RetailCrm\ApiClient::V5,
    'test'
);

$statuses = array(
    'new' => 'Новый',
    'delivery-group' => 'Доставка',
    'assembling-group' => 'Комплектация',
    'complete' => 'Выполнен',
);

$filter = array(
    'sites' => array('test')
);
$getStatuses = null;
try {
    foreach ($statuses as $status => $name){
        $filter['extendedStatus'] = $status;
        $response = $client->request->ordersList($filter);
        $getStatuses = $client->request->statusesList()->statuses;
        $dataOrders[$status] = $response;
    }
    //$status = null;
    //$name = null;

} catch (\RetailCrm\Exception\CurlException $e) {
    echo "Connection error: " . $e->getMessage();
}

if ($response->isSuccessful()) {
    $xls = new PHPExcel();
    $index = 0;
    foreach ($statuses as $status => $name){
        $xls->createSheet($index);
        $xls->setActiveSheetIndex($index);
        $sheet = $xls->getActiveSheet();
        $sheet->setTitle($name);
        $sheet->setCellValue("A1", "Заказы со статусом $name");

        $sheet->setCellValue("A2", "ФИО");
        $sheet->setCellValue("B2", "Телефон");
        $sheet->setCellValue("C2", "Адрес");
        $sheet->setCellValue("D2", "статус заказа");
        $sheet->setCellValue("E2", "Перезвонить клиенту");

        $sheet->mergeCells('A1:E1');
        $sheet->getStyle('A1')->getAlignment()->setHorizontal(
            PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        //var_dump($dataOrders);
        $stringNum = 3;
        foreach ($dataOrders[$status]->orders as $order){

            $sheet->setCellValue("A$stringNum", $order['firstName']);
            $sheet->setCellValue("B$stringNum", $order['phone']);
            $sheet->setCellValue("C$stringNum", $order['customer']['address']['text']);
            $sheet->setCellValue("D$stringNum", $getStatuses[$order['status']]['name']);
            $sheet->setCellValue("E$stringNum", $order['customFields']['recall']);
            $stringNum++;
        }
        $index++;
        $stringNum = 3;
    }
    //

    // Выводим HTTP-заголовки
    //header('Content-Type: text/html; charset=utf-8');
    header ( "Expires: Mon, 1 Apr 1974 05:00:00 GMT" );
    header ( "Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT" );
    header ( "Cache-Control: no-cache, must-revalidate" );
    header ( "Pragma: no-cache" );
    header ( "Content-type: application/vnd.ms-excel" );
    header ( "Content-Disposition: attachment; filename=retailcrm.xls" );

// Выводим содержимое файла
    $objWriter = new PHPExcel_Writer_Excel5($xls);
    $objWriter->save('php://output');
    $xls = null;

} else {
    echo sprintf(
        "Error: [HTTP-code %s] %s",
        $response->getStatusCode(),
        $response->getErrorMsg()
    );

    // error details
    //if (isset($response['errors'])) {
    //    print_r($response['errors']);
    //}
}

?>