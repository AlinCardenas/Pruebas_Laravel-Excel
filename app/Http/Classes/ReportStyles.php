<?php
namespace App\Http\Classes;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class ReportStyles 
{
    /* Estilos primer encabezado */
    public $styleFirstHead = [
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => [
                'rgb' => '538DD5',
            ],
        ],
        'font' => [
            'bold' => true,
            'color' => [
                'rgb' => 'FFFFFF',
            ],
            'size' => 16,
        ],
        'alignment' => [
            'wrapText' => true,
            'vertical' => Alignment::VERTICAL_CENTER,
            'horizontal' => Alignment::HORIZONTAL_CENTER,
        ],
    ];

    public function assignSizes($sheet){
        $columnSize = [
            'G' => 15,
            'H' => 15,
            'I' => 15,
            'J' => 15,
            'O' => 15,
            'P' => 15,
            'S' => 15,
            'T' => 15,
            'U' => 15,
            'V' => 15,
            'L' => 25,
            'K' => 25,
            'M' => 25,
            'R' => 25,
            'B' => 35,
            'F' => 35,
        ];

        foreach ($columnSize as $column => $width) {
            $sheet->getColumnDimension($column)->setWidth($width);
        }
    }

    public function assignHeadColor($sheet){
        $sheet->getRowDimension(2)->setRowHeight(41);
        $styleArray = [
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => '4F81BD',
                ],
            ],
            'font' => [
                'bold' => true,
                'color' => [
                    'rgb' => 'FFFFFF',
                ],
                'size' => 11,
            ],
            'alignment' => [
                'wrapText' => true,
                'vertical' => Alignment::VERTICAL_CENTER,
                'horizontal' => Alignment::HORIZONTAL_CENTER,
            ],
        ];

        /* Color de los encabezados */
        $sheet->getStyle('A2:V2')->applyFromArray($styleArray);
        $sheet->getStyle('G2:I2')->applyFromArray(array_merge($styleArray, [
            'fill' => [
                'startColor' => [
                    'rgb' => '9BBB59',
                ],
            ],
            
        ]));
        $sheet->getStyle('J2')->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'C0504D',
                ],
            ],
        ]);
        $sheet->getStyle('K2:N2')->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'F79646',
                ],
            ],
        ]);

    }

    public function assignColor($sheet){
        /* Sacarar el numero de filas */
        $highestRow = $sheet->getHighestRow();

        /* Asignar colores a las celdas normales */
        $rangeColors = [
            'A3:F' => 'DCE6F1',
            'G3:I' => 'EBF1DE',
            'J3:J' => 'F2DCDB',
            'K3:N' => 'FDE9D9',
            'O3:V' => 'DCE6F1',
        ];

        foreach ($rangeColors as $range => $color) {
                $range .= $highestRow;
                $sheet->getStyle($range)->applyFromArray([
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => [
                        'rgb' => $color,
                    ],  
                ],
            ]);
        } 
    }

    public function assignColorLastCell($sheet){
        /* Sacarar el numero de filas */
        $highestRow = $sheet->getHighestRow();
        /* Agregar una columna abajo */
        $rangeColors = 'A' . ($highestRow + 1) . ':F' . ($highestRow + 1);
        $rangeColors2 = 'G' . ($highestRow + 1) . ':I' . ($highestRow + 1);
        $rangeColors3 = 'J' . ($highestRow + 1) . ':J' . ($highestRow + 1);
        $rangeColors4 = 'K' . ($highestRow + 1) . ':N' . ($highestRow + 1);
        $rangeColors5 = 'O' . ($highestRow + 1) . ':V' . ($highestRow + 1);
        
        // Agregar una nueva fila al final de la tabla con colores
        $sheet->insertNewRowBefore($highestRow + 1);
        $sheet->getStyle($rangeColors)->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => '366092',
                ],
            ],
        ]);
        $sheet->getStyle($rangeColors2)->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => '76933C',
                ],
            ],
        ]);
        $sheet->getStyle($rangeColors3)->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => '963634',
                ],
            ],
        ]);
        $sheet->getStyle($rangeColors4)->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'E26B0A',
                ],
            ],
        ]);
        $sheet->getStyle($rangeColors5)->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => '366092',
                ],
            ],
        ]);

        return $highestRow;
    }

    public function assignColorCell($highestRow, $sheet){

        $rangeColors = [
            'A3:F' => 'DCE6F1',
            'G3:I' => 'EBF1DE',
            'J3:J' => 'F2DCDB',
            'K3:N' => 'FDE9D9',
            'O3:V' => 'DCE6F1',
        ];

        foreach ($rangeColors as $range => $color) {
                $range .= $highestRow;
                $sheet->getStyle($range)->applyFromArray([
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => [
                        'rgb' => $color,
                    ],  
                ],
            ]);
        }
    }
}

?>