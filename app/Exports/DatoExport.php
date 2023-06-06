<?php

namespace App\Exports;

use App\Models\dato;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class DatoExport implements FromCollection, WithHeadings,WithStyles

{
    /**
    * @return \Illuminate\Support\Collection
    */
    public function styles(Worksheet $sheet)
    { 
        // Configuración de barra arriba del encabezado
        $sheet->mergeCells('A1:V1');
        $sheet->getRowDimension(1)->setRowHeight(60);
        $sheet->getStyle('A1:V1')->applyFromArray([
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
        ]);
        
        /* Tamaño de las celdas en los encabezados */
        $columnWidths = [
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
        foreach ($columnWidths as $column => $width) {
            $sheet->getColumnDimension($column)->setWidth($width);
        }

        // Color de los encabezados
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

        /* Sacarar el numero de filas */
        $highestRow = $sheet->getHighestRow();
        /* Sacar letra de ultima columan */
        $columnCount = $sheet->getHighestColumn();


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

        /* $highestColumn = $sheet->getHighestColumn();
        $sumRange = 'H3:' . $highestColumn . '3';

        $sumFormula = '=SUM(' . $sumRange . ')';
        $sheet->setCellValue('H' . ($highestRow + 1), $sumFormula); */

        
        // Agrega la fórmula de suma a la celda A4




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


        /* $rangeBorderColors = [
            'A3:F3' => '95B3D7',
            'G3:I3' => 'C4D79B',
            'K3:N3' => 'DA9694',
            'O3:V3' => '95B3D7',
        ]; */
        
        

    }
}
    public function headings(): array
    {
        return [
            ['REPORTE DE AVANCE DE RESULTADOS 
            FECHA 00/00/00'],
            [
                'Cliente',
                'Nombre',
                'Status recargas',
                'Status servicios',
                'Origen',
                'Ejecutivo asignado',
                "Cantidad de PDV's en toda la red",
                'Cantidad de activos en timepo aire en toda la red',
                "PDV's activos con ventas mayores a $2,000",
                'Cantidad de inactivos en tiempo aire en toda la red',
                'Compras de red (Monto pagado)',
                'Compras de red (Monto abonado)',
                'Compras de red (Monto red)',
                'Venta de red',
                'Recargas realizadas RED',
                'Pagos de servicios RED',
                'Pines RED',
                '¿COMPRA SALDO POR TRANSFERENCIA ELECTRÓNCIA?',
                'ACTIVACIONES AT&T',
                'ACTIVACIONES UNEFON',
                'ACTIVACIONES BIENCEL',
             "PDV's red que activan biencel",
            ]
        ];

    }
   
    public function collection()
    {
        return dato::all();
    }
}
