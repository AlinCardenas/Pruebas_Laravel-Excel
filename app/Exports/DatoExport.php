<?php

namespace App\Exports;

use App\Models\dato;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;





class DatoExport implements FromCollection, WithHeadings,WithStyles

{
    /**
    * @return \Illuminate\Support\Collection
    */
    public function styles(Worksheet $sheet)
    {  

        
        /* Configuracion de barra arriba del encabezado */
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
                'horizontal' =>Alignment::HORIZONTAL_CENTER,
            ],
            
        ]);

      /*   $styleArray = [
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '0000FF'], // Código de color azul en formato RGB
                ],
            ],
        ];
        
        $sheet->getStyle('A1:V1')->applyFromArray($styleArray); */
        
        
        /* Color de los encabezados */
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
                'horizontal' =>Alignment::HORIZONTAL_CENTER,
            ],
        ];

        /* Tamaño de las celdas en los encabezados */
        $columns = ['G','H','I','J','O','P','S','T','U','v'];
        $width = 15;
        foreach ($columns as $column) {
            $sheet->getColumnDimension($column)->setWidth($width);
        } 
        $columns = ['L','K','M','R'];
        $width = 25;
        foreach ($columns as $column) {
            $sheet->getColumnDimension($column)->setWidth($width);
        } 
        $columns = ['B','F'];
        $width = 35;
        foreach ($columns as $column) {
            $sheet->getColumnDimension($column)->setWidth($width);
        } 
        /* Celdas mas pequeñas */
       /*  $columns = ['G','K', 'L', 'M', 'O', 'P', 'Q', 'S', 'T', 'U', 'V','B','R', 'H', 'J', 'I','F'];
        $width = 30;
        foreach ($columns as $column) {
            $sheet->getColumnDimension($column)->setWidth($width);
        }

 */

        /* Color de las columnas de abajo */
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
        $sheet->getStyle('O2:V2')->applyFromArray($styleArray);

        /* Sacarar el numero de celdas */
        $highestRow = $sheet->getHighestRow();
        $columnCount = $sheet->getHighestColumn();


        /* Agregar una columna abajo */
        $rangeColors = 'A' . ($highestRow + 1) . ':' . $columnCount . ($highestRow + 1);

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

        $rangeBorderColors = [
            'A3:F' => '95B3D7',
            'G3:I' => 'C4D79B',
            'J3:J' => 'DA9694',
            'K3:N' => 'FABF8F',
            'O3:V' => '95B3D7',
        ];

        /* $styleArray = [
            'borders' => [
                'bottom' => ['borderStyle' => Border::BORDER_THIN,'color' => ['rgb' => $color]],
            ],
        ];

        $sheet->getStyle('A3:F3')->applyFromArray($styleArray); */

        
        foreach ($rangeBorderColors as $range => $color) {
            $range .= $highestRow;
            $sheet->getStyle($range)->applyFromArray([
                'borders' => [
                    'bottom' => [
                        'borderStyle' => Border::BORDER_THICK,
                    ]
                ],
            ]);
        }
    }
/* 
    'borders' => [
        'top' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => [
                'rgb' => $color,
            ],
        ],
    ], */

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
