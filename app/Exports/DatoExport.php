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




class DatoExport implements FromCollection, WithHeadings,WithStyles

{
    /**
    * @return \Illuminate\Support\Collection
    */

    public function styles(Worksheet $sheet)
    {
        
        /* Color de los encabezados */
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
            ],
            'alignment' => [
                'wrapText' => true,
                'vertical' => Alignment::VERTICAL_BOTTOM,
            ],
        ];
        $sheet->getStyle('A1:V1')->applyFromArray($styleArray);
        $sheet->getStyle('G1:I1')->applyFromArray(array_merge($styleArray, [
            'fill' => [
                'startColor' => [
                    'rgb' => '9BBB59',
                ],
            ],
        ]));
        $sheet->getStyle('J1')->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'C0504D',
                ],
            ],
        ]);
        $sheet->getStyle('K1:N1')->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'F79646',
                ],
            ],
        ]);
        $sheet->getStyle('O1:V1')->applyFromArray($styleArray);

        /* Agregar color a las celdas de abajo */
        $highestRow = $sheet->getHighestRow();
        $columnCount = $sheet->getHighestColumn();
        $rangeColors = 'A' . ($highestRow + 1) . ':' . $columnCount . ($highestRow + 1);

         // Agregar una nueva fila al final de la tabla con colores
        $sheet->insertNewRowBefore($highestRow + 1);
            
        $sheet->getStyle($rangeColors)->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                'rgb' => 'DCE6F1',
            ],
        ],
    ]);
        $rangeColors = [
            'A2:F' => 'DCE6F1',
            'G2:I' => 'EBF1DE',
            'J2:J' => 'F2DCDB',
            'K2:N' => 'FDE9D9',
            'O2:V' => 'DCE6F1',
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
    public function headings(): array
    {
        return [
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
        ];
    }
    
    public function collection()
    {
        return dato::all();
    }
}
