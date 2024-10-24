<?php

namespace App\Imports;

use App\Models\PackingList;
use Maatwebsite\Excel\Concerns\ToModel;
use Maatwebsite\Excel\Concerns\WithColumnLimit;
use Maatwebsite\Excel\Concerns\WithHeadingRow;
use Carbon\Carbon;
use PhpOffice\PhpSpreadsheet\Shared\Date;


class PackingListImport implements ToModel,WithHeadingRow, WithColumnLimit
{
    protected $batch;

    public function __construct($batch)
    {
        $this->batch = $batch;

    }

    public function model(array $row)
    {
        $packing_list = $this->existCheck($row['factory_po'],$row['po'],$this->convertDate($row['crd']),$row['ship_mode'],$row['destination']);

        if($packing_list === null){

            $this->insertPackingList($row);
        }else{

            $no_of_sizes = $packing_list->pl_no_of_sizes;
            $name_of_sizes = $packing_list->pl_name_of_sizes;
            $name_of_size_codes = $packing_list->pl_name_of_size_codes;
            $quantities = $packing_list->pl_quantities;

            $name_of_size = ',' . $row['size'];
            $name_of_size_code = ',' . $row['size_code'];
            $quantity = ',' . $row['quantity'];

            $new_no_of_sizes = $no_of_sizes + 1;
            $new_name_of_sizes = $name_of_sizes . $name_of_size;
            $new_name_of_size_codes = $name_of_size_codes . $name_of_size_code;
            $new_quantities = $quantities . $quantity;

            $packing_list->update([
                'pl_no_of_sizes' => $new_no_of_sizes,
                'pl_name_of_sizes' => $new_name_of_sizes,
                'pl_name_of_size_codes' => $new_name_of_size_codes,
                'pl_quantities' => $new_quantities,
            ]);

//            $this->insertPackingListContent($packing_list->id,$row['size_code'],$row['size'],$row['quantity']);
        }
        return [];
    }

    protected function insertPackingList($row){

        $packing_list = [
            'pl_md' => $row['md'],
            'pl_brand' => 'JACKWOLFSKIN',
            'pl_factory_po' => $row['factory_po'],
            'pl_po' => $row['po'],
            'pl_style_code' => $row['style_code'],
            'pl_color_code' => $row['color_code'],
            'pl_style_desc' => $row['description'],
            'pl_color_desc' => $row['color'],
            'pl_crd' => $this->convertDate($row['crd']),
            'pl_ship_mode' => $row['ship_mode'],
            'pl_destination' => $row['destination'],
            'pl_customer_warehouse' => $row['customer_warehouse'],
            'pl_no_of_sizes' => 1,
            'pl_name_of_sizes' => $row['size'],
            'pl_name_of_size_codes' => $row['size_code'],
            'pl_quantities' => $row['quantity'],
            'pl_status' => 'DRAFT',
            'pl_mcq_basis' => $row['mcq_basis'],
            'batch_id' => $this->batch,

        ];

        $packing_list = PackingList::create($packing_list);

//        $this->insertPackingListContent($packing_list->id,$row['size_code'],$row['size'],$row['quantity']);

    }

   public function endColumn(): string
   {
      return 'O';
   }

    protected function convertDate($crd){
        return $crd = Carbon::instance(Date::excelToDateTimeObject($crd));
    }

    private function existCheck($factory_po, $po, $crd, $ship_mode, $destination)
    {
        $packing_list = PackingList::where([
            ['pl_factory_po',$factory_po],
            ['pl_po',$po],
            ['pl_crd',$crd],
            ['pl_ship_mode',$ship_mode],
            ['pl_destination',$destination],
            ['batch_id',$this->batch],
        ])->first();

        return $packing_list;

    }

    private function insertPackingListContent($id, $size_code, $size, $quantity)
    {


    }

}