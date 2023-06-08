<?php
namespace App\Http\Classes;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use Maatwebsite\Excel\Facades\Excel;

class ReportStyles 
{
    public $styleFirstHead;
    /* Estilos primer encabezado */
    public function styleHead($sheet){
        $sheet->mergeCells('A1:V1');
        $sheet->getStyle('A1')->getFont()->setSize(16);
        $sheet->getRowDimension(1)->setRowHeight(60);
        $this->styleFirstHead = [
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
            ],
            'alignment' => [
                'wrapText' => true,
                'vertical' => Alignment::VERTICAL_CENTER,
                'horizontal' => Alignment::HORIZONTAL_CENTER,
            ],
        ];
        $sheet->getStyle('A1:V1')->applyFromArray( $this->styleFirstHead);
    }
    /* Tamaño de las celdas */
    public function assignSizes($sheet)
    {
        $columnSize = [
            'B' => 45,
            'E' => 12.86,
            'F' => 28,
            'G' => 12.86,
            'H' => 12.86,
            'I' => 12.86,
            'J' => 9.43,
            'K' => 14.86,
            'L' => 14.86,
            'M' => 14.86,
            'N' => 14.86,
            'O' => 12.71,
            'P' => 12.71,
            'R' => 12.71,
            'S' => 12.71,
            'T' => 12.71,
            'U' => 12.71,
            'V' => 12.71, 
        ];
        foreach ($columnSize as $column => $width) {
            $sheet->getColumnDimension($column)->setWidth($width);
        }
    }
    /* Color de encabezados */
    public function assignHeadColor($sheet)
    {
        $sheet->getRowDimension(2)->setRowHeight(41);
        /* Color de los encabezados */
        $sheet->getStyle('A2:V2')->applyFromArray(array_merge($this->styleFirstHead, [
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => '4F81BD',
                ],
            ],
        ]));
        $sheet->getStyle('G2:I2')->applyFromArray(array_merge($this->styleFirstHead, [
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => '9BBB59',
                ],
            ],
        ]));
        $sheet->getStyle('J2')->applyFromArray(array_merge($this->styleFirstHead, [
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'C0504D',
                ],
            ],
        ]));
        $sheet->getStyle('K2:N2')->applyFromArray(array_merge($this->styleFirstHead, [
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'F79646',
                ],
            ],
        ]));
    }
    /* Color de celdas */
    public function assignColorCell($highestRow, $sheet)
    {
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

    public function assignColorLastCell($sheet)
    {
        $highestRow = $sheet->getHighestRow();

        $rangeColors = 'A' . ($highestRow + 1) . ':F' . ($highestRow + 1);
        $rangeColors2 = 'G' . ($highestRow + 1) . ':I' . ($highestRow + 1);
        $rangeColors3 = 'J' . ($highestRow + 1) . ':J' . ($highestRow + 1);
        $rangeColors4 = 'K' . ($highestRow + 1) . ':N' . ($highestRow + 1);
        $rangeColors5 = 'O' . ($highestRow + 1) . ':V' . ($highestRow + 1);
        
        $sheet->insertNewRowBefore($highestRow + 1);

        $colorStyles = [
            $rangeColors => '366092',
            $rangeColors2 => '76933C',
            $rangeColors3 => '963634',
            $rangeColors4 => 'E26B0A',
            $rangeColors5 => '366092',
        ];

        foreach ($colorStyles as $range => $color) {
            $sheet->getStyle($range)->applyFromArray([
                'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'startColor' => [
                        'rgb' => $color,
                    ],
                ],
            ]);
        }

        return $highestRow;
    }

    public function assignLastText($sheet, $highestRow)
    {
        $texCell = 'D' . ($highestRow + 1);
        $sheet->setCellValue($texCell, 'Totales');
        $styleArray = [
            'alignment' => $this->styleFirstHead['alignment'],
            'font' => $this->styleFirstHead['font'],
        ];
        $sheet->getStyle($texCell)->applyFromArray($styleArray);
    }

    public function addSymbol($sheet)
    {
        $columns = ['K', 'L', 'M', 'N'];
        $highestRow = $sheet->getHighestRow();

        $styleArray = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
            ],
            'numberFormat' => [
                'formatCode' => '$#,##0.00',
            ],
        ];

        foreach ($columns as $column) {
            $range = $column . '3:' . $column . $highestRow;
            $sheet->getStyle($range)->applyFromArray($styleArray);
        }
    }

    public function alignment($sheet, $highestRow)
    {
        $ranges = [
            'A3:A' . $highestRow,
            'G3:I' . $highestRow,
            'J3:J' . $highestRow,
            'O3:V' . $highestRow,
        ];
        $styleArray['alignment']['horizontal'] = Alignment::HORIZONTAL_CENTER;
        foreach ($ranges as $range) {
            $sheet->getStyle($range)->applyFromArray($styleArray);
        }
    }

    public function cal($totals, $sheet, $highestRow)
    {
        $cellData = [
            'G' => 'totalRed',
            'H' => 'totalActivos',
            'I' => 'totalActivos2000',
            'J' => 'totalInactivos',
            'K' => 'totalPagado',
            'L' => 'totalAbonado',
            'M' => 'totalAbonoAred',
            'N' => 'ventasMensuales',
            'O' => 'totalRecargas',
            'P' => 'totalServicios',
            'Q' => 'totalPines',
            'R' => 'transferenciaS',
            'S' => 'activacionesAtt',
            'T' => 'activacionesUnefon',
            'U' => 'activacionesBiencel',
            'V' => 'usuariosBiencel'
        ];

        foreach ($cellData as $column => $dataKey) {
            $texCell = $column . ($highestRow + 1);
            $sheet->setCellValue($texCell, $totals[$dataKey]);
            $styleArray = [
                'alignment' => $this->styleFirstHead['alignment'],
                'font' => $this->styleFirstHead['font'],
            ];
            $sheet->getStyle($texCell)->applyFromArray($styleArray);
        }
    }
}
?>