<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up(): void
    {
        Schema::create('datosprueba', function (Blueprint $table) {
            $table->id();
            $table->string('cliente');
            $table->string('nombre');
            $table->string('status_recarga');
            $table->string('status_servicios');
            $table->string('origen');
            $table->string('ejecutivo_asignado');
            $table->integer('cantidad_pdv');
            $table->integer('cantidad_activos');
            $table->integer('cantidad_inactivos');
            $table->integer('PDV');
            $table->integer('cantidad_inactivos_tiempo');
            $table->integer('compras_pagado');
            $table->integer('compras_abonado');
            $table->integer('compras_red');
            $table->integer('venta_red');
            $table->integer('recargas_red');
            $table->integer('pagos_red');
            $table->integer('pines_red');
            $table->integer('saldo_transf');
            $table->string('activaciones_at');
            $table->string('activaciones_unefon');
            $table->string('activaciones_biencel');
            $table->string('red_activa');
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('datosprueba');
    }
};
