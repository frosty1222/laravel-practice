<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Response;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
class TestController extends Controller
{
    public function index(){
        return view('welcome');
    }
    public function exportData(){
        return [
            [
                "name"=>"lo van dong",
                "phone"=>"09342343224",
                "status"=>"single",
                "gender"=>"male",
                "address"=>[
                    [
                        "alley"=>"xuan phuong nam tu liem",
                        "house Number"=>19,
                        "seria"=>"adsadasdas",
                        "seria1"=>"adsadasdas",
                        "seria2"=>"adsadasdas",
                        "seria3"=>"adsadasdas",
                        "seria4"=>"adsadasdas",
                    ],
                    [
                        "alley"=>"xuan phuong nam tu liem",
                        "house Number"=>19,
                        "seria"=>"adsadasdas",
                        "seria1"=>"adsadasdas",
                        "seria2"=>"adsadasdas",
                        "seria3"=>"adsadasdas",
                        "seria4"=>"adsadasdas",
                    ],
                    [
                        "alley"=>"xuan phuong nam tu liem",
                        "house Number"=>19,
                        "seria"=>"adsadasdas",
                        "seria1"=>"adsadasdas",
                        "seria2"=>"adsadasdas",
                        "seria3"=>"adsadasdas",
                        "seria4"=>"adsadasdas",
                    ],
                ],
                "nationality"=>"Viet nam"
            ],
            [
                "name"=>"lo van dong",
                "phone"=>"09342343224",
                "status"=>"single",
                "gender"=>"male",
                "address"=>[
                    [
                        "alley"=>"xuan phuong nam tu liem",
                        "house Number"=>19,
                        "seria"=>"adsadasdas",
                        "seria1"=>"adsadasdas",
                        "seria2"=>"adsadasdas",
                        "seria3"=>"adsadasdas",
                        "seria4"=>"adsadasdas",
                    ]
                ],
                "nationality"=>"Viet nam"
            ],
            [
                "name"=>"lo van dong",
                "phone"=>"09342343224",
                "status"=>"single",
                "gender"=>"male",
                "address"=>[
                    [
                    "alley"=>"xuan phuong nam tu liem",
                    "house Number"=>19,
                    "seria"=>"adsadasdas",
                    "seria1"=>"adsadasdas",
                    "seria2"=>"adsadasdas",
                    "seria3"=>"adsadasdas",
                    "seria4"=>"adsadasdas",
                    ]
                ],
                "nationality"=>"Viet nam"
            ],
            [
                "name"=>"lo van dong",
                "phone"=>"09342343224",
                "status"=>"single",
                "gender"=>"male",
                "address"=>[
                    "alley"=>"xuan phuong nam tu liem",
                    "house Number"=>19
                ],
                "nationality"=>"Viet nam"
            ],
            [
                "name"=>"lo van dong",
                "phone"=>"09342343224",
                "status"=>"single",
                "gender"=>"male",
                "address"=>[
                    "alley"=>"xuan phuong nam tu liem",
                    "house Number"=>19
                ],
                "nationality"=>"Viet nam"
            ],
            [
                "name"=>"lo van dong",
                "phone"=>"09342343224",
                "status"=>"single",
                "gender"=>"male",
                "address"=>[
                    "alley"=>"xuan phuong nam tu liem",
                    "house Number"=>19
                ],
                "nationality"=>"Viet nam"
            ],
        ];
    }
    public function exportExcel(){
        $data = $this->exportData();
       // Create a new Spreadsheet instance
       $spreadsheet = new Spreadsheet();
       $sheet = $spreadsheet->getActiveSheet();

       // Set headers based on the array keys
       $headers = array_keys($data[0]);
       $col = 'A';
       foreach ($headers as $header) {
           $sheet->setCellValue($col . '1', $header);
           $col++;
       }

       // Add data
       $row = 2;
       foreach ($data as $item) {
           $col = 'A';
           foreach ($item as $key => $value) {
               if ($key === "address") {
                   if (is_array($value)) {
                       // If $value is an array, concatenate its values with line breaks
                       $addressString = '';
                       foreach ($value as $addressItem) {
                           if (is_array($addressItem)) {
                               $addressString .= implode(', ', $addressItem) . "\n";
                           } else {
                               $addressString .= $addressItem . "\n";
                           }
                       }
                       $sheet->setCellValue($col . $row, rtrim($addressString, "\n"));
                   } else {
                       // If it's not an array, set it as is
                       $sheet->setCellValue($col . $row, $value);
                   }
               } else {
                   // For other fields, set the value as is
                   $sheet->setCellValue($col . $row, $value);
               }
               $col++;
           }
           $row++;
       }

       // Create a new Excel writer
       $writer = new Xlsx($spreadsheet);

       // Save the Excel file to a specific path
       $excelFileName = 'data_export.xlsx';
       $writer->save($excelFileName);

       // Provide a download link for the user
       return Response::download($excelFileName, 'data_export.xlsx');
    }
}
