<?php

namespace Database\Factories;

use Illuminate\Database\Eloquent\Factories\Factory;

/**
 * @extends \Illuminate\Database\Eloquent\Factories\Factory<\App\Models\Datoprueba>
 */
class DatosPruebaFactory extends Factory
{
    /**
     * Define the model's default state.
     *
     * @return array<string, mixed>
     */
    public function definition(): array
    {
        return [
            'cliente' => $this->faker->word,
            'nombre' => $this->faker->word,
            'status_recarga' => $this->faker->word,
            'status_servicios' => $this->faker->word,
            'origen' => $this->faker->word,
            'ejecutivo_asignado' => $this->faker->word,
            'cantidad_pdv' => $this->faker->numberBetween(1,1000),
            'cantidad_activos' => $this->faker->numberBetween(1,50),
            'cantidad_inactivos' => $this->faker->numberBetween(1,50),
            'PDV' => $this->faker->numberBetween(1,50),
            'cantidad_inactivos_tiempo' => $this->faker->numberBetween(1,50),
            'compras_pagado' => $this->faker->numberBetween(1,50),
            'compras_abonado' => $this->faker->numberBetween(1,50),
            'compras_red' => $this->faker->numberBetween(1,50),
            'venta_red' => $this->faker->numberBetween(1,50),
            'recargas_red' => $this->faker->numberBetween(1,50),
            'pagos_red' => $this->faker->numberBetween(1,50),
            'pines_red' => $this->faker->numberBetween(1,50),
            'saldo_transf' => $this->faker->numberBetween(1,50),
            'activaciones_at' => $this->faker->word,
            'activaciones_unefon' => $this->faker->word,
            'activaciones_biencel' => $this->faker->word,
            'red_activa' => $this->faker->word,
            
        ];
    }
}
