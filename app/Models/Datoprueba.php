<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Datoprueba extends Model
{

    use HasFactory;
    protected $fillable=[
        'cliente',
        'nombre',
        'status_recarga',
        'status_servicios',
        'origen',
        'ejecutivo_asignado',
        'cantidad_pdv',
        'cantidad_activos',
        'PDV',
        'cantidad_inactivos_tiempo',
        'compras_pagado',
        'compras_abonado',
        'compras_red',
        'venta_red',
        'recargas_red',
        'pagos_red',
        'pines_red',
        'saldo_transf',
        'activaciones_at',
        'activaciones_unefon',
        'activaciones_biencel',
        'red_activa',
    ];
}
