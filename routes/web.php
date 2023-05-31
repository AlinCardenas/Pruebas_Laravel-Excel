<?php

use App\Exports\DatoExport;
use Illuminate\Support\Facades\Route;
use Maatwebsite\Excel\Concerns\FromCollection;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
*/

Route::get('/', function () {
    return Excel::download(new DatoExport, 'users.xlsx');
    return view('welcome');
});
