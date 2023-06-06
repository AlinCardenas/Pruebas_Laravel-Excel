<?php

namespace App\Exports;
use App\Http\Classes\ReportStyles;
use App\Models\dato;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class DatoExport implements FromCollection, WithHeadings,WithStyles
{
    /**
    * @return \Illuminate\Support\Collection
    */
    public function styles(Worksheet $sheet)
    
        $objStyle = new ReportStyles();
        // Configuración primer encabezado
        $sheet->mergeCells('A1:V1');
        $sheet->getRowDimension(1)->setRowHeight(60);
        $sheet->getStyle('A1:V1')->applyFromArray(
            $objStyle->styleFirstHead
        );
        /* Tamaño de las celdas en los encabezados */
        $objStyle->assignSizes($sheet);
        // Color de los encabezados
        $objStyle->assignHeadColor($sheet);
        $highestRow = $objStyle->assignColorLastCell($sheet);
        $objStyle->assignColorCell($highestRow, $sheet); 
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
