<?php

namespace App\Exports;

use App\Models\Carton;
use App\Models\Style;
use App\Models\user;
use Carbon\Carbon;
use Illuminate\Support\Facades\Storage;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithDrawings;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;
use Maatwebsite\Excel\Events\BeforeSheet;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class PackingListExport implements FromCollection,WithTitle,WithEvents,WithDrawings
{

    protected $packing_list;
    protected $column_letter;
    protected $size_list;
    protected $pl_sizes_sort;

    protected $pl_quantities;
    protected $pl_size_letters;
    protected $pl_weight;
    protected $pl_first_carton;
    protected $pl_second_carton;
    protected $pl_first_mcq;
    protected $pl_second_mcq;
    protected $pl_third_carton;
    protected $pl_third_mcq;
    protected $pl_balance_quantity;
    protected $table_content_row_start;
    protected $table_second_content_row_start;
    protected $pl_weight_cell;

    protected $pl_first_carton_weight;
    protected $pl_second_carton_weight;
    protected $pl_third_carton_weight;



    protected $pl_carton_number;
    protected $pl_no_size_mcq;
    protected $pl_next_ctn_mcq;
    protected $pl_carton_list;

    protected $numberSeparator;
    protected $numberSeparatorDecimal;

    public function __construct($packing_list)
    {
        $this->packing_list = $packing_list;

//        dd($this->packing_list);
        $this->column_letter = ['G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB'
                                ,'AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU'];
        $this->size_list = ['XS','S','M','L','XL','2XL','3XL'];

        $cartons = Carton::all();

        if(count($cartons) > 0){
            $this->pl_first_carton_weight = $cartons[0]->ctn_weight;
        }else{
            $this->pl_first_carton_weight = 1.3;
        }
        if(count($cartons) > 1){
            $this->pl_second_carton_weight = $cartons[1]->ctn_weight;
        }else{
            $this->pl_second_carton_weight = 1.1;
        }
        if(count($cartons) > 2){
            $this->pl_third_carton_weight = $cartons[2]->ctn_weight;
        }else{
            $this->pl_third_carton_weight = 0.9;
        }

        $this->pl_third_carton_weight = 1;

        $this->pl_first_mcq = [];
        $this->pl_second_mcq = [];
        $this->pl_third_mcq = [];

        $this->numberSeparator = "#,##0";
        $this->numberSeparatorDecimal = "#,##0.00";

        $this->table_content_row_start = 12;
        $this->table_second_content_row_start = 12;

        $this->pl_carton_number = 1;
        $this->pl_balance_quantity = [];
        $this->pl_no_size_mcq = [];
        $this->pl_carton_list = [];

        $this->initiateSizeDetails();
        //initiate pl sizes sort, pl size codes , pl quantities by sizes
        //dump($this->pl_sizes_sort);
        //dd($this->pl_quantities);
    }

    public function collection()
    {
        return collect([]);
    }

    public function title(): string
    {
        return $this->packing_list['pl_style_code'] . ' ' . $this->packing_list['pl_color_desc'];
    }

    public function registerEvents(): array
    {
       return [
           BeforeSheet::class => function(BeforeSheet $event){

                //SET COLUMN WITH INITIAL VIEW
               $this->setColumnsWidth($event);
               //SET COLUMN WITH INITIAL VIEW
               //SET TABLE HEADER
               $this->exportTableHeader($event);
               //SET TABLE HEADER
               //SET PL HEADER
               $this->exportPLHeader($event);
               //SET PL HEADER
               //BORDER
               $event->sheet->mergeCells('A8:'.($this->column_letter[$this->packing_list['pl_no_of_sizes']+8]).'8')->getStyle('A8:'.($this->column_letter[$this->packing_list['pl_no_of_sizes']+8]).'8')
                   ->applyFromArray(['borders' => ['bottom' => ['borderStyle' => Border::BORDER_MEDIUM,],],]);
               //BORDER
               //DISPLAY GARMENT AND SIZES INFOR START AF
                $this->displaySizeMCQDetails($event);
               //DISPLAY GARMENT AND SIZES INFOR START AF
               //DISPLAY TABLE CONTENT
               $this->displayTableContent($event);
               //DISPLAY TABLE CONTENT
               //BORDER
               $event->sheet->mergeCells('A'.($this->table_second_content_row_start+1).':'.($this->column_letter[$this->packing_list['pl_no_of_sizes']+8]).($this->table_second_content_row_start+1))
                   ->getStyle('A'.($this->table_second_content_row_start+1).':'.($this->column_letter[$this->packing_list['pl_no_of_sizes']+8]).($this->table_second_content_row_start+1))
                   ->applyFromArray(['borders' => ['bottom' => ['borderStyle' => Border::BORDER_MEDIUM,],],]);
               //BORDER
               //DISPLAY SIZE SUMMARY
                $this->displaySizeSummary($event);
               //DISPLAY SIZE SUMMARY
               //DISPLAY carton SUMMARY
                $this->displayCartonSummary($event);
               //DISPLAY carton SUMMARY
               //DISPLAY PACKING METHOD
               $this->displayPackingMethod($event);
               //DISPLAY PACKING METHOD
               //DISPLAY CARTON MARKING
               $this->displayCartonMarking($event);
               //DISPLAY CARTON MARKING

           },
           AfterSheet::class => function(AfterSheet $event){

               $event->sheet
                   ->getPageSetup()
                   ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE)
                   ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4)
                   ->setScale(65)
                  ;
               $event->sheet->getPageMargins()->setHeader(0.1);
               $event->sheet->getPageMargins()->setTop(0.1);
               $event->sheet->getPageMargins()->setLeft(0.2);
               $event->sheet->getPageMargins()->setBottom(0.1);
               $event->sheet->getPageMargins()->setRight(0.2);
               $event->sheet->getPageMargins()->setFooter(0.1);
           }
       ];
    }

    private function setColumnsWidth($event)
    {
        $event->sheet->getColumnDimension('A')->setWidth(7);
        $event->sheet->getColumnDimension('B')->setWidth(7);
        $event->sheet->getColumnDimension('C')->setWidth(15);
        $event->sheet->getColumnDimension('D')->setWidth(20);
        $event->sheet->getColumnDimension('E')->setWidth(20);
        $event->sheet->getColumnDimension('F')->setWidth(15);

        for($x=0;$x < $this->packing_list['pl_no_of_sizes'];$x++){
            $event->sheet->getColumnDimension($this->column_letter[$x])->setWidth(7);
        }
        $event->sheet->getColumnDimension($this->column_letter[$this->packing_list['pl_no_of_sizes']])->setWidth(9);
        $event->sheet->getColumnDimension($this->column_letter[$this->packing_list['pl_no_of_sizes']+1])->setWidth(9);
        $event->sheet->getColumnDimension($this->column_letter[$this->packing_list['pl_no_of_sizes']+2])->setWidth(10);
        $event->sheet->getColumnDimension($this->column_letter[$this->packing_list['pl_no_of_sizes']+3])->setWidth(9);
        $event->sheet->getColumnDimension($this->column_letter[$this->packing_list['pl_no_of_sizes']+4])->setWidth(10);
        $event->sheet->getColumnDimension($this->column_letter[$this->packing_list['pl_no_of_sizes']+5])->setWidth(9);
        $event->sheet->getColumnDimension($this->column_letter[$this->packing_list['pl_no_of_sizes']+6])->setWidth(10);
        $event->sheet->getColumnDimension($this->column_letter[$this->packing_list['pl_no_of_sizes']+7])->setWidth(15);
        $event->sheet->getColumnDimension($this->column_letter[$this->packing_list['pl_no_of_sizes']+8])->setWidth(10);

    }

    private function exportTableHeader($event)
    {
        $style_head = [
            'borders' => [
                //outline all
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'color' => ['argb' => '0c0d0d']
            ],
            'font' => [
                'size' => 8,
                'bold' => true,
                'color' => ['argb' => 'FFFFFF'],
            ],
        ];
        $event->sheet->mergeCells('A10:B11');
        $event->sheet->setCellValue('A10','CTN NO箱号')->getStyle('A10:B11')->applyFromArray($style_head);
        $event->sheet->mergeCells('C10:C11');
        $event->sheet->setCellValue('C10','PO')->getStyle('C10:C11')->applyFromArray($style_head);
        $event->sheet->mergeCells('D10:D11');
        $event->sheet->setCellValue('D10','Style')->getStyle('D10:D11')->applyFromArray($style_head);
        $event->sheet->mergeCells('E10:E11');
        $event->sheet->setCellValue('E10','Description')->getStyle('E10:E11')->applyFromArray($style_head);
        $event->sheet->mergeCells('F10:F11');
        $event->sheet->setCellValue('F10','Color')->getStyle('F10:F11')->applyFromArray($style_head);

        $event->sheet->mergeCells($this->column_letter[0].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']-1].'10');
        $event->sheet->setCellValue($this->column_letter[0].'10','SIZE RATIO')
            ->getStyle($this->column_letter[0].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']-1].'10')
            ->applyFromArray($style_head);
        //SIZES
        for($z=0;$z < $this->packing_list['pl_no_of_sizes'];$z++){
            $event->sheet->setCellValue($this->column_letter[$z].'11',$this->pl_sizes_sort[$z])->getStyle($this->column_letter[$z].'11')
                ->applyFromArray($style_head);
        }
        //SIZES

        $cell = $this->column_letter[$this->packing_list['pl_no_of_sizes']].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']].'11';
        $event->sheet->mergeCells($cell);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']].'10','配比数量QTY')->getStyle($cell)->applyFromArray($style_head);

        $cell = $this->column_letter[$this->packing_list['pl_no_of_sizes']+1].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+1].'11';
        $event->sheet->mergeCells($cell);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+1].'10','TTL CARTON')->getStyle($cell)->applyFromArray($style_head);

        $cell = $this->column_letter[$this->packing_list['pl_no_of_sizes']+2].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+2].'11';
        $event->sheet->mergeCells($cell);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+2].'10','TTL QTY (PCS)')->getStyle($cell)->applyFromArray($style_head);

        $cell = $this->column_letter[$this->packing_list['pl_no_of_sizes']+3].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+3].'11';
        $event->sheet->mergeCells($cell);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+3].'10','NW KGS')->getStyle($cell)->applyFromArray($style_head);

        $cell = $this->column_letter[$this->packing_list['pl_no_of_sizes']+4].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+4].'11';
        $event->sheet->mergeCells($cell);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+4].'10','TTL NW')->getStyle($cell)->applyFromArray($style_head);

        $cell = $this->column_letter[$this->packing_list['pl_no_of_sizes']+5].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+5].'11';
        $event->sheet->mergeCells($cell);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+5].'10','GW KGS')->getStyle($cell)->applyFromArray($style_head);

        $cell = $this->column_letter[$this->packing_list['pl_no_of_sizes']+6].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+6].'11';
        $event->sheet->mergeCells($cell);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+6].'10','TTL GW')->getStyle($cell)->applyFromArray($style_head);

        $cell = $this->column_letter[$this->packing_list['pl_no_of_sizes']+7].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+7].'11';
        $event->sheet->mergeCells($cell);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+7].'10','Carton Size(CM) (LxWxH)')->getStyle($cell)->applyFromArray($style_head);

        $cell = $this->column_letter[$this->packing_list['pl_no_of_sizes']+8].'10:'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+8].'11';
        $event->sheet->mergeCells($cell);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+8].'10','CBM')->getStyle($cell)->applyFromArray($style_head);

    }



    private function exportPLHeader($event)
    {
        $style = [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'font' => [
                'size' => 15,
                'bold' => true,
            ],
        ];
        $event->sheet->mergeCells('G1:N1');
        $event->sheet->setCellValue('G1','PACKING LIST')->getStyle('G1:N1')->applyFromArray($style);
        $event->sheet->mergeCells('G2:N2');
        $event->sheet->setCellValue('G2',$this->packing_list['pl_factory_po'])->getStyle('G2:N2')->applyFromArray($style)->applyFromArray(['font' => ['size' => 10,],]);
        $event->sheet->mergeCells('G3:N3');
        $event->sheet->setCellValue('G3',$this->packing_list['pl_ship_mode'])->getStyle('G3:N3')->applyFromArray($style)->applyFromArray(['font' => ['size' => 10,],]);

        //DISPLAY PHOTO
        $worksheet = $event->sheet->getDelegate();
        $this->setBrandLogo($worksheet);
        //DISPLAY PHOTO
        $style = [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'font' => [
                'size' => 10,
                'bold' => true,
            ],
        ];
        $event->sheet->mergeCells('A2:B2');
        $event->sheet->mergeCells('A3:B3');
        $event->sheet->mergeCells('A4:B4');
        $event->sheet->mergeCells('A5:B5');
        $event->sheet->mergeCells('A6:B6');
        $event->sheet->mergeCells('A7:B7');

        $event->sheet->setCellValue('A2','Status: ')->getStyle('A2')->applyFromArray($style);
        $event->sheet->setCellValue('A3','MD: ')->getStyle('A3')->applyFromArray($style);
        $event->sheet->setCellValue('A4','Print Date: ')->getStyle('A4')->applyFromArray($style);
        $event->sheet->setCellValue('A5','CRD: ')->getStyle('A5')->applyFromArray($style);
        $event->sheet->setCellValue('A6','Customer: ')->getStyle('A6')->applyFromArray($style);
        $event->sheet->setCellValue('A7','Destination: ')->getStyle('A7')->applyFromArray($style);


        $event->sheet->setCellValue('C2',$this->packing_list['pl_status']);
        $event->sheet->setCellValue('C3',$this->packing_list['pl_md']);
        $event->sheet->setCellValue('C4','=now()')->getStyle('C4')->applyFromArray([ 'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        ],]);
        $event->sheet
            ->getStyle('C4')
            ->getNumberFormat()
            ->setFormatCode('yyyy-mm-dd');

        $event->sheet->setCellValue('C5',$this->packing_list['pl_crd']);

        $event->sheet->setCellValue('C6',$this->packing_list['pl_customer_warehouse']);

        $event->sheet->setCellValue('C7',$this->packing_list['pl_destination']);

    }

    public function drawings()
    {

        $drawings = new Drawing();
        $drawings->setName('Horizon');
        $drawings->setDescription('Horizon Logo');
        $drawings->setPath(public_path('\\storage\\images\\horizon-logo.png'));
        $drawings->setHeight(25);
        $drawings->setCoordinates('A1');
        return [$drawings];

    }

    private function setBrandLogo($worksheet)
    {
        $drawings = new Drawing();
        $drawings->setName('BLAUER');
        $drawings->setDescription('BLAUER Logo');
        $drawings->setPath(public_path('\\storage\\images\\bl-logo.png'));
        $drawings->setHeight(25);
        $drawings->setCoordinates('C1');
        $drawings->setWorksheet($worksheet);

    }

    private function initiateSizeDetails()
    {
        $exp_name_of_sizes = explode(',',$this->packing_list['pl_name_of_sizes']);
        $exp_quantities = explode(',',$this->packing_list['pl_quantities']);

        for($y=0;$y<=$this->packing_list['pl_no_of_sizes'];$y++){
            $this->pl_sizes_sort[array_search($exp_name_of_sizes[$y], $this->size_list)] = $exp_name_of_sizes[$y];
            $this->pl_quantities[$exp_name_of_sizes[$y]] = $exp_quantities[$y];
            $this->pl_size_letters[$exp_name_of_sizes[$y]] = $this->column_letter[26+$y];
        }
        $this->pl_sizes_sort = array_values(collect($this->pl_sizes_sort)->sortKeys()->toArray());
    }

    private function displaySizeMCQDetails($event)
    {
        $event->sheet->setCellValue($this->column_letter[25].'1','Style Sizes Details');
        $event->sheet->setCellValue($this->column_letter[25].'2','Sizes');
        $event->sheet->setCellValue($this->column_letter[25].'3','Weight');
        $event->sheet->setCellValue($this->column_letter[25].'4','1st Carton');
        $event->sheet->setCellValue($this->column_letter[25].'5','1st MCQ');
        $event->sheet->setCellValue($this->column_letter[25].'6','2nd Carton');
        $event->sheet->setCellValue($this->column_letter[25].'7','2nd MCQ');
        $event->sheet->setCellValue($this->column_letter[25].'8','3rd Carton');
        $event->sheet->setCellValue($this->column_letter[25].'9','3rd MCQ');

        $style = Style::where('style_code',$this->packing_list['pl_mcq_basis'])->first();

        $ac = 0;

        //26 was AF in column letter
        for($a=26;$a < (26+$this->packing_list['pl_no_of_sizes']);$a++){
            $event->sheet->setCellValue($this->column_letter[$a].'2',$this->pl_sizes_sort[$ac]);

            if(isset($style->mcq_contents)){
                if(count($style->mcq_contents->where('style_size',$this->pl_sizes_sort[$ac])) !== 0){
                    //CHECK IF HAVE STYLE SIZE DETAILS
                    $weight = array_values($style->mcq_contents->where('style_size',$this->pl_sizes_sort[$ac])->toArray())[0]['style_weight'];
                    $event->sheet->setCellValue($this->column_letter[$a].'3',$weight);
                    $this->pl_weight[$ac] = $weight;
                    $this->pl_weight_cell[$ac] = $this->column_letter[$a].'3';

                    $first_mcq = $style->mcq_contents->where('style_size',$this->pl_sizes_sort[$ac])->max('mcq');
                    $first_carton = array_values($style->mcq_contents->where('style_size',$this->pl_sizes_sort[$ac])->where('mcq',$first_mcq)->toArray())[0]['carton_measurement'];
                    $this->pl_first_mcq[$ac] = $first_mcq;
                    $this->pl_first_carton[$ac] = $first_carton;
                    $event->sheet->setCellValue($this->column_letter[$a].'4',$first_carton);
                    $event->sheet->setCellValue($this->column_letter[$a].'5',$first_mcq);



                    if(count($style->mcq_contents->where('style_size',$this->pl_sizes_sort[$ac])) >1){
                        $second_mcq =  (int)array_values($style->mcq_contents->where('style_size',$this->pl_sizes_sort[$ac])->sortByDesc('mcq')->toArray())[1]['mcq'];
                        $second_carton = array_values($style->mcq_contents->where('style_size',$this->pl_sizes_sort[$ac])->sortByDesc('mcq')->toArray())[1]['carton_measurement'];
                        $this->pl_second_mcq[$ac] = $second_mcq;
                        $this->pl_second_carton[$ac] = $second_carton;
                        $event->sheet->setCellValue($this->column_letter[$a].'6',$second_carton);
                        $event->sheet->setCellValue($this->column_letter[$a].'7',$second_mcq);

                    }

                    if(count($style->mcq_contents->where('style_size',$this->pl_sizes_sort[$ac])) >2){
                        $third_mcq =  (int)array_values($style->mcq_contents->where('style_size',$this->pl_sizes_sort[$ac])->sortByDesc('mcq')->toArray())[2]['mcq'];
                        $third_carton = array_values($style->mcq_contents->where('style_size',$this->pl_sizes_sort[$ac])->sortByDesc('mcq')->toArray())[2]['carton_measurement'];
                        $this->pl_third_mcq[$ac] = $third_mcq;
                        $this->pl_third_carton[$ac] = $third_carton;
                        $event->sheet->setCellValue($this->column_letter[$a].'8',$third_carton);
                        $event->sheet->setCellValue($this->column_letter[$a].'9',$third_mcq);
                    }


                }
            }
            $ac++;
        }


    }

    private function displayTableContent( $event)
    {
        $style = [
            'borders' => [
                //outline all
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'font' => [
                'size' => 10,
            ],
        ];

//        dump($this->pl_weight);
//        dump($this->pl_first_carton);
//        dump($this->pl_first_mcq);
//        dump($this->pl_second_carton);
//        dump($this->pl_second_mcq);
//        dd($this->pl_sizes_sort);

        //PUTTING FIRST QUANTITIES BY MCQ
        $this->insertFirstQuantity($event);
        //PUTTING FIRST QUANTITIES BY MCQ
        //PUTTING 2ND QUANTITIES MIX PACKS
        if($this->pl_balance_quantity !== null){
            $this->insertBalanceQuantity($event);
        }
        //PUTTING 2ND QUANTITIES MIX PACKS
        //PUTTING SMALL CARTON

        if($this->pl_next_ctn_mcq !== null){
//            dd($this->pl_next_ctn_mcq);
            $this->insertNextCartonMcq($event);
        }
        //

        // NEED TO SECOND ROW TOTAL SO PUT THE STYLE EB DESCRIPTION AND COLOR HERE
        $event->sheet->mergeCells('C12:C'.($this->table_second_content_row_start-1));
        $event->sheet->setCellValue('C12',$this->packing_list['pl_po'])
            ->getStyle('C12:C'.($this->table_second_content_row_start-1))->applyFromArray($style);
        $event->sheet->mergeCells('D12:D'.($this->table_second_content_row_start-1));
        $event->sheet->setCellValue('D12',$this->packing_list['pl_style_code'])
            ->getStyle('D12:D'.($this->table_second_content_row_start-1))->applyFromArray($style);
        $event->sheet->mergeCells('E12:E'.($this->table_second_content_row_start-1));
        $event->sheet->setCellValue('E12',$this->packing_list['pl_style_desc'])
            ->getStyle('E12:E'.($this->table_second_content_row_start-1))->applyFromArray($style);
        $event->sheet->mergeCells('F12:F'.($this->table_second_content_row_start-1));
        $event->sheet->setCellValue('F12',($this->packing_list['pl_color_code'].'_'.$this->packing_list['pl_color_desc']))
            ->getStyle('F12:F'.($this->table_second_content_row_start-1))->applyFromArray($style);

        //DISPLAY TOTAL SUM OF THE QUANTITY DETAILS
        $style_foot = [
            'borders' => [
                //outline all
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'color' => ['argb' => '0c0d0d']
            ],
            'font' => [
                'size' => 12,
                'bold' => true,
                'color' => ['argb' => 'FFFFFF'],
            ],
        ];
        //ttl carton TOTAL
        $ttl_ctn = '=SUM('.$this->column_letter[$this->packing_list['pl_no_of_sizes']+1].$this->table_content_row_start.
            ':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_second_content_row_start-1).')';
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_second_content_row_start), $ttl_ctn)
            ->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_second_content_row_start))->applyFromArray($style_foot);
        //ttl qty pcs
        $ttl_qty_pcs = '=SUM('.$this->column_letter[$this->packing_list['pl_no_of_sizes']+2].$this->table_content_row_start.
            ':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+2].($this->table_second_content_row_start-1).')';
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+2].($this->table_second_content_row_start), $ttl_qty_pcs)
            ->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+2].($this->table_second_content_row_start))->applyFromArray($style_foot);
        //NW total
        $ttl_nw ='=SUM('.$this->column_letter[$this->packing_list['pl_no_of_sizes']+3].$this->table_content_row_start.
            ':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_second_content_row_start-1).')';
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_second_content_row_start), $ttl_nw)
            ->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_second_content_row_start))->applyFromArray($style_foot);
        //NW total total
        $ttl_nw_ttl ='=SUM('.$this->column_letter[$this->packing_list['pl_no_of_sizes']+4].$this->table_content_row_start.
            ':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+4].($this->table_second_content_row_start-1).')';
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+4].($this->table_second_content_row_start), $ttl_nw_ttl)
            ->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+4].($this->table_second_content_row_start))->applyFromArray($style_foot);
        //GW total
        $ttl_gw ='=SUM('.$this->column_letter[$this->packing_list['pl_no_of_sizes']+5].$this->table_content_row_start.
            ':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+5].($this->table_second_content_row_start-1).')';
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+5].($this->table_second_content_row_start), $ttl_gw)
            ->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+5].($this->table_second_content_row_start))->applyFromArray($style_foot);
        //GW total
        $ttl_gw_ttl ='=SUM('.$this->column_letter[$this->packing_list['pl_no_of_sizes']+6].$this->table_content_row_start.
            ':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+6].($this->table_second_content_row_start-1).')';
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+6].($this->table_second_content_row_start), $ttl_gw_ttl)
            ->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+6].($this->table_second_content_row_start))->applyFromArray($style_foot);
        //GW total
        $cbm ='=SUM('.$this->column_letter[$this->packing_list['pl_no_of_sizes']+8].$this->table_content_row_start.
            ':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+8].($this->table_second_content_row_start-1).')';
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+8].($this->table_second_content_row_start), $cbm)
            ->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+8].($this->table_second_content_row_start))->applyFromArray($style_foot);

        $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+7].($this->table_second_content_row_start))->applyFromArray($style_foot);
        $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']].($this->table_second_content_row_start))->applyFromArray($style_foot);
        $event->sheet->mergeCells('A'.$this->table_second_content_row_start.':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']-1].$this->table_second_content_row_start);
        $event->sheet->getStyle('A'.$this->table_second_content_row_start.':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']-1].$this->table_second_content_row_start)
            ->applyFromArray($style_foot);
        //DISPLAY TOTAL SUM OF THE QUANTITY DETAILS

        //DISPLAY THE SIZE THAT NO MCQ
        if(count($this->pl_no_size_mcq) > 0){
            $this->insertNoSizeMcq($event);
        }
        //DISPLAY THE SIZE THAT NO MCQ

        //PUT BORDER TO CONTENT
        $style_brd = [
            'borders' => [
                //outline all
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
        ];


        for($brd = $this->table_content_row_start;$brd <= $this->table_second_content_row_start; $brd++){
            if($brd < $this->table_second_content_row_start){
                $event->sheet->getStyle('A'.$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparator);
                $event->sheet->getStyle('B'.$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparator);
                $event->sheet->getStyle('D'.$brd)->applyFromArray($style_brd);

                for($brds = 0; $brds < $this->packing_list['pl_no_of_sizes'];$brds++){
                    $event->sheet->getStyle($this->column_letter[$brds].$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparator);
                }
            }
            $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']].$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparator);
            $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+1].$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparator);
            $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+2].$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparator);
            $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+3].$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparatorDecimal);
            $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+4].$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparatorDecimal);
            $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+5].$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparatorDecimal);
            $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+6].$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparatorDecimal);
            $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+7].$brd)->applyFromArray($style_brd);
            $event->sheet->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+8].$brd)->applyFromArray($style_brd)->getNumberFormat()->setFormatCode($this->numberSeparatorDecimal);
        }

        //PUT BORDER TO CONTENT

    }

    private function    insertFirstQuantity($event)
    {
        $frc = 0;
        for($ss = 0 ;$ss < $this->packing_list['pl_no_of_sizes'];$ss++){
            $size = $this->pl_sizes_sort[$ss];
            $quantity = $this->pl_quantities[$size];

            if(array_key_exists($ss,$this->pl_first_mcq)){
                //check if the mcq or carton are have if dont have skip this line no else

                if($quantity >= $this->pl_first_mcq[$ss]){

                    //first size mcq
                    $event->sheet->setCellValue($this->column_letter[$ss].($this->table_content_row_start+$frc),
                        $this->pl_first_mcq[$ss]);
                    //QTY
                    $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']].($this->table_content_row_start+$frc),
                        '='.$this->column_letter[$ss].($this->table_content_row_start+$frc));
                    //ttl carton
                    $ttl_carton = (int)($quantity/$this->pl_first_mcq[$ss]);
                    $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_content_row_start+$frc),
                        $ttl_carton);
                    //ttl qty pcs
                    $ttl_qty_pcs = $this->pl_first_mcq[$ss]*$ttl_carton;
                    $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+2].($this->table_content_row_start+$frc),
                        $ttl_qty_pcs);
                    //NW
                    $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_content_row_start+$frc),
                        '=' .
                        $this->column_letter[$this->packing_list['pl_no_of_sizes']].($this->table_content_row_start+$frc) .
                        '*' .
                        $this->pl_weight_cell[$ss]);
                    //TOTAL NW
                    $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+4].($this->table_content_row_start+$frc),
                        '=' .
                        $this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_content_row_start+$frc) .
                        '*' .
                        $this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_content_row_start+$frc));
                    //GW
                    $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+5].($this->table_content_row_start+$frc),
                        '=' .
                        $this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_content_row_start+$frc) .
                        '+' .
                        $this->pl_first_carton_weight);
                    //TOTAL GW
                    $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+6].($this->table_content_row_start+$frc),
                        '=' .
                        $this->column_letter[$this->packing_list['pl_no_of_sizes']+5].($this->table_content_row_start+$frc) .
                        '*' .
                        $this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_content_row_start+$frc));
                    //CARTON MEASUREMENT
                    $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+7].($this->table_content_row_start+$frc),
                        $this->pl_first_carton[$ss]);
                    if(!in_array($this->pl_first_carton[$ss],$this->pl_carton_list)){
                        array_push($this->pl_carton_list,$this->pl_first_carton[$ss]);
                    }
                    //CBM
                    $carton_explode = explode('*',$this->pl_first_carton[$ss]);
                    $cbm = ((float)$carton_explode[0]/100) * ((float)$carton_explode[1]/100) * ((float)$carton_explode[2]/100) *
                        $ttl_carton;
                    $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+8].($this->table_content_row_start+$frc),
                        $cbm);

                    //CARTON NUMBER
                    $carton_number_two = $this->pl_carton_number + ($ttl_carton-1);
                    $event->sheet->setCellValue('A'.($this->table_content_row_start+$frc), $this->pl_carton_number);
                    $event->sheet->setCellValue('B'.($this->table_content_row_start+$frc), $carton_number_two);
                    $this->pl_carton_number = $this->pl_carton_number + $ttl_carton;

                    //CHECK IF HAVE BALANCE QUANTITY
                    $balance_qty = $quantity - $ttl_qty_pcs;
                    if($balance_qty > 0){
                        $this->pl_balance_quantity[$size] = (int)$balance_qty;
                    }

                    //ADD ROW CONTENT END
                    $this->table_second_content_row_start++;

                    //row number
                    $frc++;

                }else if($quantity < $this->pl_first_mcq[$ss]){
                    //if less than mcq put already to quantities left
                    $this->pl_balance_quantity[$size] = (int)$quantity;
                }

            }else{
                //print last row this size no mcq
                $this->pl_no_size_mcq[$size] = $quantity;
            }

        }
    }

    private function insertBalanceQuantity($event)
    {
//        dd($this->pl_balance_quantity);
//                dd($this->table_second_content_row_start);
        //FIRST SEPARATE BY HOW MANY FIRST MCQ CAN BE TOGETHER IN THE SIZE BALANCE QUANTITY
        $ctrl = 0;
        $separate_qty= [];
        $separate_size= [];
        $separate_sort= [];



        for($bq = 0;$bq < 10; $bq++){
            $separate_sort[$bq] = 0;
            if(count($this->pl_balance_quantity) > 0){
                //if have balance infinite the loop
                $ctrl = 0;

                $qty_left = 0;
                foreach(array_reverse($this->pl_balance_quantity) as $key => $bal_qty){

                        if($qty_left === 0){
                            $qty_left = $this->pl_first_mcq[array_search($key,$this->pl_sizes_sort)];
                        }

                        if($bal_qty <= $qty_left){
                            $qty_left = $qty_left - $bal_qty;

                            if($qty_left > 0){
                                $separate_qty[$bq][$ctrl] = $bal_qty;
                                $separate_size[$bq][$ctrl] = $key;
                                $separate_sort[$bq] = $separate_sort[$bq] + $bal_qty;
                                unset($this->pl_balance_quantity[$key]);
                            }else if($qty_left === 0){
                                $separate_qty[$bq][$ctrl] = $bal_qty;
                                $separate_size[$bq][$ctrl] = $key;
                                $separate_sort[$bq] = $separate_sort[$bq] + $bal_qty;
                                unset($this->pl_balance_quantity[$key]);
                                break;
                            }

                        }

                    $ctrl++;
                }

            }else{
                //if no have more balance separated kill the loop
                $bq = $bq + 20;
            }

        }

        //remove the last array because its add to array = 0 no use
        array_pop($separate_sort);

//        dump($separate_size);
//        dump($separate_qty);
//        dd(count($separate_qty));
        $count_separate_size = count($separate_size);

        for($cbq = 0; $cbq < $count_separate_size;$cbq++){
            //CHECK IF SINGLE AND CAN PUT IN SMALL CARTON
            if(count($separate_size[$cbq]) === 1){
                //CHECK IF FIT TO THE 2ND CARTON MCQ
                $qty = (int)array_values($separate_qty[$cbq])[0];
                $size = array_values($separate_size[$cbq])[0];
                $second_mcq = 0;

                if(array_key_exists(array_search($size,$this->pl_sizes_sort),$this->pl_second_mcq)){
                    $second_mcq = (int)$this->pl_second_mcq[array_search($size,$this->pl_sizes_sort)];
                }

                if($qty <= $second_mcq){
                    $this->pl_next_ctn_mcq[$size] = $qty;
                    unset($separate_qty[$cbq]);
                    unset($separate_size[$cbq]);
                    unset($separate_sort[$cbq]);
                }
            }

        }
//        dd($this->pl_next_ctn_mcq);
        //separate qty and separate size can print now first mcq
//        dump($separate_qty);
//        dump($separate_size);
        //small ctn mcq was going to small carton because of the two quantity even add cannot go at big carton

        //DISPLAY 2ND SEPARATE QTY

//        dd('please check if the balance quantity combine can put in 2nd carton');
//make sort bt the sum of array combine


        $separate_sort = collect($separate_sort)->sort()->reverse();
        foreach($separate_sort as $sp => $sp_array){

            $quantities = array_values($separate_qty[$sp]);
            $sizes = array_values($separate_size[$sp]);


            //THIS LOOP IS FOR ONLY DATA FOR 1 PRINT ONLY LIKE CBM

            $row_sum_qty = 0;
            for($spp = 0;$spp < count($quantities); $spp++){
                //THIS LOOP IS FOR ONLY DATA FOR SIZES HAVE 2 OR MORE VALUES

                $cell_size = (int)array_search($sizes[$spp], $this->pl_sizes_sort);

                //mix size mcq
                $event->sheet->setCellValue($this->column_letter[$cell_size].($this->table_second_content_row_start),
                    $quantities[$spp]);

                $event->sheet->setCellValue($this->column_letter[26+$cell_size].($this->table_second_content_row_start),
                    '='.$this->column_letter[$cell_size].($this->table_second_content_row_start).
                    '*'.$this->column_letter[26+$cell_size].'3');
                //number 3 was row for weight


                //GET ROW SUM
                $row_sum_qty = $row_sum_qty + $quantities[$spp];

                //THIS LOOP IS FOR ONLY DATA FOR SIZES HAVE 2 OR MORE VALUES
            }

            //THIS LOOP IS FOR ONLY DATA FOR 1 PRINT ONLY LIKE CBM

            //QTY
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']].($this->table_second_content_row_start),
                '=SUM('.$this->column_letter[0].$this->table_second_content_row_start.':'
                .$this->column_letter[$this->packing_list['pl_no_of_sizes']-1].$this->table_second_content_row_start.')');
            //ttl carton
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_second_content_row_start),
                1);
            //ttl qty pcs
            $ttl_qty_pcs = '='.$this->column_letter[$this->packing_list['pl_no_of_sizes']].($this->table_second_content_row_start).'*1';
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+2].($this->table_second_content_row_start),
                $ttl_qty_pcs);
            //NW
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_second_content_row_start),
                '=SUM(' .
                $this->column_letter[26+0].($this->table_second_content_row_start) .
                ':' .
                $this->column_letter[26+$this->packing_list['pl_no_of_sizes']-1].($this->table_second_content_row_start) . ')');
            //TOTAL NW
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+4].($this->table_second_content_row_start),
                '=' .
                $this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_second_content_row_start) .
                '*' .
                $this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_second_content_row_start));

            // 0 mean the bigger size
            $size_number = (int)array_search($sizes[0], $this->pl_sizes_sort);
            if(isset($this->pl_second_mcq[$size_number])){
                if($row_sum_qty <= $this->pl_second_mcq[$size_number]){
                    $carton_weight = $this->pl_second_carton_weight;
                    $carton = $this->pl_second_carton[$size_number];
                }else{
                    $carton_weight = $this->pl_first_carton_weight;
                    $carton = $this->pl_first_carton[$size_number];
                }
            }else{
                $carton_weight = $this->pl_first_carton_weight;
                $carton = $this->pl_first_carton[$size_number];
            }
            //GW
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+5].($this->table_second_content_row_start),
                '=' .
                $this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_second_content_row_start) .
                '+' .
                $carton_weight);
            //TOTAL GW
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+6].($this->table_second_content_row_start),
                '=' .
                $this->column_letter[$this->packing_list['pl_no_of_sizes']+5].($this->table_second_content_row_start) .
                '*' .
                $this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_second_content_row_start));

            //CARTON MEASUREMENT
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+7].($this->table_second_content_row_start),
                $carton);
            if(!in_array($carton,$this->pl_carton_list)){
                array_push($this->pl_carton_list,$carton);
            }
            //CBM
            $carton_explode = explode('*',$carton);
            $cbm = ((float)$carton_explode[0]/100) * ((float)$carton_explode[1]/100) * ((float)$carton_explode[2]/100) * 1;
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+8].($this->table_second_content_row_start),
                $cbm);



            //CARTON NUMBER
            $carton_number_two = $this->pl_carton_number;
            $event->sheet->setCellValue('A'.($this->table_second_content_row_start), $this->pl_carton_number);
            $event->sheet->setCellValue('B'.($this->table_second_content_row_start), $carton_number_two);
            $this->pl_carton_number = $this->pl_carton_number + 1;


            $this->table_second_content_row_start++;
        }

        //DISPLAY 2ND SEPARATE QTY
    }

    private function insertNextCartonMcq($event)
    {
        //THIS FUNCTION ARE FOR THE BALANCE QUANTITY THAT GO TO SMALL CARTON
        foreach($this->pl_next_ctn_mcq as $size => $qty){
            $cell_size = (int)array_search($size, $this->pl_sizes_sort);

            //mix size mcq
            $event->sheet->setCellValue($this->column_letter[$cell_size].($this->table_second_content_row_start),
                $qty);
            //QTY
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']].($this->table_second_content_row_start),
                $qty);
            //ttl carton
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_second_content_row_start),
                1);
            //DISPLAY NW PER SIZE
            $event->sheet->setCellValue($this->column_letter[26+$cell_size].($this->table_second_content_row_start),
                '='.$this->column_letter[$cell_size].($this->table_second_content_row_start).
                '*'.$this->column_letter[26+$cell_size].'3');
            //ttl qty pcs
            $ttl_qty_pcs = '='.$this->column_letter[$this->packing_list['pl_no_of_sizes']].($this->table_second_content_row_start).'*1';
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+2].($this->table_second_content_row_start),
                $ttl_qty_pcs);
            //NW
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_second_content_row_start),
                '=SUM(' .
                $this->column_letter[26+0].($this->table_second_content_row_start) .
                ':' .
                $this->column_letter[26+$this->packing_list['pl_no_of_sizes']-1].($this->table_second_content_row_start) . ')');
            //TOTAL NW
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+4].($this->table_second_content_row_start),
                '=' .
                $this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_second_content_row_start) .
                '*' .
                $this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_second_content_row_start));

            $size_number = (int)array_search($size, $this->pl_sizes_sort);
//            dd($qty . ' ' . $this->pl_third_mcq[$size_number]);
            if(isset($this->pl_third_mcq[$size_number])){
                if($qty <= $this->pl_third_mcq[$size_number]){

                    $carton_weight = $this->pl_third_carton_weight;
                    $carton = $this->pl_third_carton[$size_number];
                }else{
                    $carton_weight = $this->pl_second_carton_weight;
                    $carton = $this->pl_second_carton[$size_number];
                }
            }else{
                if(isset($this->pl_second_mcq[$size_number])) {
                    $carton_weight = $this->pl_second_carton_weight;
                    $carton = $this->pl_second_carton[$size_number];
                }else{
                    $carton_weight = $this->pl_first_carton_weight;
                    $carton = $this->pl_first_carton[$size_number];
                }
            }
            //GW
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+5].($this->table_second_content_row_start),
                '=' .
                $this->column_letter[$this->packing_list['pl_no_of_sizes']+3].($this->table_second_content_row_start) .
                '+' .
                $carton_weight);
            //TOTAL GW
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+6].($this->table_second_content_row_start),
                '=' .
                $this->column_letter[$this->packing_list['pl_no_of_sizes']+5].($this->table_second_content_row_start) .
                '*' .
                $this->column_letter[$this->packing_list['pl_no_of_sizes']+1].($this->table_second_content_row_start));
            //CARTON MEASUREMENT
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+7].($this->table_second_content_row_start),
                $carton);
            if(!in_array($carton,$this->pl_carton_list)){
                array_push($this->pl_carton_list,$carton);
            }
            //CBM
            $carton_explode = explode('*',$carton);
            $cbm = ((float)$carton_explode[0]/100) * ((float)$carton_explode[1]/100) * ((float)$carton_explode[2]/100) * 1;
            $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+8].($this->table_second_content_row_start),
                $cbm);

            //CARTON NUMBER
            $carton_number_two = $this->pl_carton_number;
            $event->sheet->setCellValue('A'.($this->table_second_content_row_start), $this->pl_carton_number);
            $event->sheet->setCellValue('B'.($this->table_second_content_row_start), $carton_number_two);
            $this->pl_carton_number = $this->pl_carton_number + 1;

            $this->table_second_content_row_start++;

        }

    }

    private function insertNoSizeMcq($event)
    {
        $style = [
            'font' => [
                'size' => 10,
                'bold' => true,
                'color' => ['argb' => 'ff0000'],
            ],
        ];

        $event->sheet->setCellValue('S1','WARNING!!!')->getStyle('S1')->applyFromArray($style);
        $event->sheet->setCellValue('S2','NO MCQ!!!')->getStyle('S2')->applyFromArray($style);
        $nmc = 1;
        foreach($this->pl_no_size_mcq as $size => $qty){
            $event->sheet->setCellValue('T'.$nmc,$size . '-' . $qty)->getStyle('T'.$nmc)->applyFromArray($style);
            $nmc++;
        }
    }

    private function displaySizeSummary($event)
    {
        $size_summary_row_start = $this->table_second_content_row_start+3;
        $style = [
            'borders' => [
                //outline all
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'font' => [
                'size' => 10,
                'bold' => true,
            ],
        ];
//        dd($this->pl_quantities);
        $event->sheet->setCellValue('F'.$size_summary_row_start,'Size Summary')->getStyle('F'.$size_summary_row_start)->applyFromArray($style);
        //SIZES AND VALUE
        for($z=0;$z < $this->packing_list['pl_no_of_sizes'];$z++){
            $event->sheet->setCellValue($this->column_letter[$z].$size_summary_row_start,$this->pl_sizes_sort[$z])
                ->getStyle($this->column_letter[$z].$size_summary_row_start)
                ->applyFromArray($style);
            $event->sheet->setCellValue($this->column_letter[$z].($size_summary_row_start+1),
                $this->pl_quantities[$this->pl_sizes_sort[$z]])
                ->getStyle($this->column_letter[$z].($size_summary_row_start+1))
                ->applyFromArray($style)->getNumberFormat()->setFormatCode($this->numberSeparator);
        }
        //SIZES AND VALUE
        $event->sheet->setCellValue('F'.($size_summary_row_start+1),$this->packing_list['pl_color_desc'])
            ->getStyle('F'.($size_summary_row_start+1))->applyFromArray($style);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']].($size_summary_row_start),
            'TOTAL')->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']].($size_summary_row_start))->applyFromArray($style);
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']].($size_summary_row_start+1),
            '=SUM('.$this->column_letter[0].($size_summary_row_start+1).':'.
            $this->column_letter[$this->packing_list['pl_no_of_sizes']-1].($size_summary_row_start+1).')')
            ->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']].($size_summary_row_start+1))->applyFromArray($style)
            ->getNumberFormat()->setFormatCode($this->numberSeparator);
    }

    private function displayCartonSummary( $event)
    {
        $carton_summary_row_start = $this->table_second_content_row_start+6;
        $style = [
            'borders' => [
                //outline all
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'font' => [
                'size' => 10,
                'bold' => true,
            ],
        ];

        $event->sheet->setCellValue('F'.$carton_summary_row_start,'Carton Summary')
            ->getStyle('F'.$carton_summary_row_start)->applyFromArray($style);
        $event->sheet->setCellValue('G'.$carton_summary_row_start,'QTY')
            ->getStyle('G'.$carton_summary_row_start)->applyFromArray($style);

        for($cl = 1; $cl <= count($this->pl_carton_list);$cl++){
            $event->sheet->setCellValue('F'.($carton_summary_row_start+$cl),$this->pl_carton_list[$cl-1])
                ->getStyle('F'.($carton_summary_row_start+$cl))->applyFromArray($style);
            $carton_qty = '=SUMIF($'.($this->column_letter[$this->packing_list['pl_no_of_sizes']+7]).'$'.$this->table_content_row_start.':'.
                '$'.($this->column_letter[$this->packing_list['pl_no_of_sizes']+7]).'$'.($this->table_second_content_row_start-1).
                ',F'.($carton_summary_row_start+$cl).','.
                '$'.($this->column_letter[$this->packing_list['pl_no_of_sizes']+1]).'$'.$this->table_content_row_start.':'.
                '$'.($this->column_letter[$this->packing_list['pl_no_of_sizes']+1]).'$'.($this->table_second_content_row_start-1).')';
            $event->sheet->setCellValue('G'.($carton_summary_row_start+$cl),$carton_qty)
                ->getStyle('G'.($carton_summary_row_start+$cl))->applyFromArray($style)->getNumberFormat()->setFormatCode($this->numberSeparator);
        }
//        dd($this->pl_carton_list);
    }

    private function displayPackingMethod( $event)
    {
        $style = [
            'borders' => [
                //outline all
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'font' => [
                'size' => 10,
                'bold' => true,
            ],
        ];
        $event->sheet->mergeCells('A'.($this->table_second_content_row_start+3).':D'.($this->table_second_content_row_start+3));
        $event->sheet->setCellValue('A'.($this->table_second_content_row_start+3),'Packing Method')
            ->getStyle('A'.($this->table_second_content_row_start+3).':D'.($this->table_second_content_row_start+3))->applyFromArray($style);
        $event->sheet->mergeCells('A'.($this->table_second_content_row_start+4).':D'.($this->table_second_content_row_start+8));
        $event->sheet->getStyle('A'.($this->table_second_content_row_start+4).':D'.($this->table_second_content_row_start+8))->applyFromArray($style);
    }

    private function displayCartonMarking($event)
    {
        $style = [
            'borders' => [
                //outline all
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'font' => [
                'size' => 10,
                'bold' => true,
            ],
        ];

        $event->sheet->mergeCells($this->column_letter[$this->packing_list['pl_no_of_sizes']+2].($this->table_second_content_row_start+3).
            ':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+8].($this->table_second_content_row_start+3));
        $event->sheet->setCellValue($this->column_letter[$this->packing_list['pl_no_of_sizes']+2].($this->table_second_content_row_start+3),'Carton Mark')
            ->getStyle($this->column_letter[$this->packing_list['pl_no_of_sizes']+2].($this->table_second_content_row_start+3).
                ':'.$this->column_letter[$this->packing_list['pl_no_of_sizes']+8].($this->table_second_content_row_start+3))->applyFromArray($style);

        //DISPLAY carton mark
        $carton_mark = $event->sheet->getDelegate();
        $this->setCartonMark($carton_mark);
        //DISPLAY carton mark
    }

    private function setCartonMark($carton_mark)
    {

        $fileName = '/public/images/carton-mark/'.$this->packing_list['pl_customer_warehouse'].'.png';
        $fileNameTwo = public_path('/storage/images/carton-mark/'.$this->packing_list['pl_customer_warehouse'].'.png');
        $cell = $this->column_letter[$this->packing_list['pl_no_of_sizes']+2].($this->table_second_content_row_start+4);

        $exists = Storage::disk('local')->exists($fileName);

        if($exists){
            $drawings = new Drawing();
            $drawings->setName('Carton Mark');
            $drawings->setDescription($this->packing_list['pl_customer_warehouse']);
            $drawings->setPath($fileNameTwo);
            $drawings->setHeight(150);
            $drawings->setCoordinates($cell);
            $drawings->setWorksheet($carton_mark);
        }


    }

}
