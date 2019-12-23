<?php

namespace app\index\controller;

use think\Controller;
use PHPExcel;
use PHPExcel_IOFactory;

class Excel extends Controller
{
    /**边框可选
     * BORDER_NONE             = 'none';
     * BORDER_DASHDOT          = 'dashDot';
     * BORDER_DASHDOTDOT       = 'dashDotDot';
     * BORDER_DASHED           = 'dashed';
     * BORDER_DOTTED           = 'dotted';
     * BORDER_DOUBLE           = 'double';
     * BORDER_HAIR             = 'hair';
     * BORDER_MEDIUM           = 'medium';
     * BORDER_MEDIUMDASHDOT    = 'mediumDashDot';
     * BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot';
     * BORDER_MEDIUMDASHED     = 'mediumDashed';
     * BORDER_SLANTDASHDOT     = 'slantDashDot';
     * BORDER_THICK            = 'thick';
     * BORDER_THIN             = 'thin';
     */
    /**
     * $PHPExcel->createSheet(); //创建一个sheet
     * $PHPExcel->setActiveSheetIndex(1); //设置创建的sheet为活动工作簿
     * $PHPExcel->getActiveSheet()->setTitle('Simple') //重命名工作簿
     *
     * @method export() 导出excel或保存本地，默认导出可设置$filename路径保存本地
     * @method sheet(number $len) 生成工作簿 $len数量
     * @method filename(string $filename) 生成excel文件名
     * @method header(array $head) 表格头部 ['李四:aaa', '张三:test'];
     * 左边是头值，右边是数据键
     * @method body(array $data) 表格数据 二位数组或三维数组 二位数组一个工作簿，三维数组多个工作簿
     * @method write(string $filename) 默认浏览器输出，有值时保存服务器
     *
     */
    protected $phpExcel;
    protected $Excel2007 = 'phpoffice/phpexcel/Classes/PHPExcel/Writer/Excel2007';
    protected $Excel5 = 'phpoffice/phpexcel/Classes/PHPExcel/Writer/Excel5';
    //当前活动工作簿
    protected $ActiveSheetIndex = 0;
    // 默认1 Excel2007 否则0  Excel5
    public $type = 1;
    //文件名
    public $filename = '公司汇结.xlsx';
    //sheet工作簿数量 默认为1个
    public $sheetlen = 1;
    //表格头部信息
    public $head = [];
    //A-Z
    public $range = [];
    //sheet 名
    public $sheetname = 'sheet';

    //单元格边框粗细
    public $borderweight = \PHPExcel_Style_Border::BORDER_NONE;

    public $title = [];

    /**
     * 实例化PHPExcel
     */
    public function initialize()
    {
        $this->phpExcel = new PHPExcel();
        //设置所有单元格居中
        $this->phpExcel->getDefaultStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->int();
    }

    /**
     * 基本设置
     */
    protected function int()
    {
        //设置属性
        $this->phpExcel->getProperties()->setCreator("Maarten Balliauw")
            ->setLastModifiedBy("Maarten Balliauw")
            ->setTitle("Office 2007 XLSX Test Document")
            ->setSubject("Office 2007 XLSX Test Document")
            ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
            ->setKeywords("office 2007 openxml php")
            ->setCategory("Test result file");
        $this->range = range("A", "Z");
    }

    public function export()
    {
        $this->write();
    }

    /**
     * 额外生成工作簿
     * @param object $PHPExcel PHPExcel对象
     * @param number $len 生成sheet个数 从0开始数
     */
    public function sheet($len)
    {
        if ($len - 1 > 0) {
            for ($i = 1; $i <= $len - 1; $i++) {
                $this->phpExcel->createSheet();
            }
            $this->sheetlen = $len;
        }
        return $this;
    }


    /**
     * 设置标题
     */
    public function title($title)
    {
        $this->title = $title;
        return $this;
    }


    /**
     * 设置文件名
     */
    public function filename($filename)
    {
        $this->filename = $filename;
        return $this;
    }

    /**
     * 设置sheet名，多个用 “|”分隔
     */
    public function sheetname($sheetname)
    {
        $this->sheetname = $sheetname;
        return $this;
    }
    /**[设置单元格元素位置]*/
    private function align($align)
    {
        $a = null;
        switch($align){
            case 'general':
                $a = \PHPExcel_Style_Alignment::HORIZONTAL_GENERAL;
                break;
            case 'left':
                $a = \PHPExcel_Style_Alignment::HORIZONTAL_LEFT;
                break;
            case 'right';
                $a = \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
                break;
            case 'center':
                $a = \PHPExcel_Style_Alignment::HORIZONTAL_CENTER;
                break;
            case 'centerContinuous':
                $a = \PHPExcel_Style_Alignment::HORIZONTAL_CENTER_CONTINUOUS;
                break;
            case 'justify':
                $a = \PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY;
                break;
            case 'fill':
                $a = \PHPExcel_Style_Alignment::HORIZONTAL_FILL;
                break;
            case 'distributed':
                $a = \PHPExcel_Style_Alignment::HORIZONTAL_DISTRIBUTED;
                break;
            default:
                $a = \PHPExcel_Style_Alignment::HORIZONTAL_CENTER;
        }
        return $a;
    }
    /**
     * 设置边框粗细
     */
    public function border($border)
    {

        $style = null;
        switch ($border) {
            case 'dashDot':
                $style = \PHPExcel_Style_Border::BORDER_DASHDOT;
                break;
            case 'dashDotDot':
                $style = \PHPExcel_Style_Border::BORDER_DASHDOTDOT;
                break;
            case 'dashed':
                $style = \PHPExcel_Style_Border::BORDER_DASHED;
                break;
            case 'dotted':
                $style = \PHPExcel_Style_Border::BORDER_DOTTED;
                break;
            case 'double':
                $style = \PHPExcel_Style_Border::BORDER_DOUBLE;
                break;
            case 'hair':
                $style = \PHPExcel_Style_Border::BORDER_HAIR;
                break;
            case 'medium':
                $style = \PHPExcel_Style_Border::BORDER_MEDIUM;
                break;
            case 'mediumDashDot':
                $style = \PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT;
                break;
            case 'mediumDashDotDot':
                $style = \PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT;
                break;
            case 'mediumDashed':
                $style = \PHPExcel_Style_Border::BORDER_MEDIUMDASHED;
                break;
            case 'slantDashDot':
                $style = \PHPExcel_Style_Border::BORDER_SLANTDASHDOT;
                break;
            case 'thick':
                $style = \PHPExcel_Style_Border::BORDER_THICK;
                break;
            case 'thin':
                $style = \PHPExcel_Style_Border::BORDER_THIN;
                break;
            default:
                $style = \PHPExcel_Style_Border::BORDER_NONE;
        }
        $this->borderweight = $style;
        return $this;
    }

    /**
     * excel头部
     * @param array $head
     * [标题:字段:宽度:文本位置(left/right...)]
     * 张三:name:18:right
     */
    public function header($head)
    {
        $this->head = $head;
        return $this;

    }

    /**
     * 数据处理
     * @param object $PHPExcel PHPExcel对象
     * @param array $data 数据 3维数组
     * [
     *      0=>[
     *          0=>[],
     *          1=>[]
     *      ],
     *      1=>[
     *         0=>[],
     *         1=>[]
     *      ]
     * ]
     */
    public function body($data)
    {
        $sheetname = explode('|', $this->sheetname);
        $is_title = count($this->title);

        $countRow = $this->countRow();
        foreach ($data as $k => $v) {
            if ($this->sheetlen > 1) { //到这里说明要分sheet了
                if (!isset($v[0])) throw new \Exception('数据必须是三维数组');
                $this->phpExcel->setActiveSheetIndex($k);
                $this->ActiveSheetIndex = $k; //当前活动工作簿
                $this->setBody($k,count($v)+$countRow);
                $this->setTitle();
                $this->setHead();
                $this->phpExcel->getActiveSheet()->setTitle($sheetname[$k]); //工作簿名称
                $this->setCellValue($v);
            } else {
                if (isset($v[0])) throw new \Exception('数据必须是二维数组');
                $this->setCellValue([$v]);
                $this->phpExcel->getActiveSheet()->setTitle($sheetname[0]);
                $this->setBody(0,count($v));
                $this->setTitle();
                $this->setHead();
            }
        }
        $this->phpExcel->setActiveSheetIndex(0);
        return $this;
    }
    /**
     * 设置单元格对齐方式
     * @param string $cell 单元格 'A1'
     * @param string $align 'right'...
     */
    private function setAlign($cell,$align){
        $a = $this->align($align);
        $this->phpExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal($a);
    }
    /**
     * [设置表格边框]
     * @param int $k 当前活动sheet
     * @param int $len 数据行数
     */
    private function setBody($k=0,$len=0)
    {
        $this->phpExcel//设置单元格边框
        ->getActiveSheet($k)
            ->getStyle('A1:' . $this->range[count($this->head) - 1] . $len)
            ->applyFromArray(array(
                'borders' => array(
                    'allborders' => array( //设置全部边框
                        'style' => $this->borderweight //粗的是thick
                    ),
                ),
            ));
    }
    /**
     * [设置单元格值]
     * @param array $data 一位数组
     */
    private function setCellValue($data)
    {
        $cell_start = $this->countRow()+1;
        foreach ($data as $ks => $vs) {
            $head = $this->head;
            foreach ($head as $kh => $vh) {
                $vr = explode(':', $vh);
                $this->phpExcel->getActiveSheet()->setCellValue($this->range[$kh] . ($ks + $cell_start), $vs[$vr[1]]);
                if(isset($vr[3])) $this->setAlign($this->range[$kh] . ($ks + $cell_start),$vr[3]);
            }
        }
    }
    /**
     * [计算数据行以外的行数]
     * 包括标题，表格头
     */
    public function countRow()
    {
        $cell_start = 0;
        if(count($this->head) > 0 ) $cell_start += 1;
        if(count($this->title) > 0) $cell_start+=1;
        return $cell_start;
    }
    //设置标题信息
    private function setTitle()
    {
        if (!empty($this->title)) {
            $cell = $this->range[0] . '1:' . $this->range[count($this->head) - 1] . '1';
            $this->phpExcel->getActiveSheet()->mergeCells($cell);
            $this->phpExcel->getActiveSheet()->setCellValue($this->range[0] . '1', $this->title[$this->phpExcel->getActiveSheetIndex()]);
        }
    }

    //设置当前head 头信息
    private function setHead()
    {
        if(count($this->head) == 0) return;
        $cell = '1';
        if (!empty($this->title)) $cell = '2';
        foreach ($this->head as $k => $v) {
            $v = explode(':', $v);
            //设置头部内容
            $this->phpExcel->getActiveSheet()->setCellValue($this->range[$k] . $cell, $v[0]);
            //设置头部宽度
            if (isset($v[2])) $this->phpExcel->getActiveSheet()->getColumnDimension($this->range[$k])->setWidth($v[2]);
        }
    }

    /**
     * 写入文件
     * @param object $PHPExcel PHPExcel对象
     * @param string $filename 默认浏览器输出下载，需要本地保存文件填写文件目录
     * @param
     */
    public function write($filename = "php://output")
    {
        if ($this->type == 1) {
            $type = 'Excel2007';
        } else {
            $type = 'Excel5';
        }
        if (strpos($filename, 'php://output') !== false) {
            $this->output();
        } else {
            $root_path = app()->getRootPath();
            $file = $root_path . 'upload/excel/' . $filename;
            if (file_exists($file)) unlink($file);
            $this->createdir('excel/' . $filename);
            $filename = $file;
        }
        $objWriter = \PHPExcel_IOFactory::createWriter($this->phpExcel, $type);
        $objWriter->save($filename);
    }

    /**如果要输出浏览器*/
    private function output()
    {
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $this->filename . '"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0
    }

    /**
     * 创建文件夹
     * 指向根路径
     */
    function createdir($path, $oi = 1)
    {
        $zpath = explode('/', $path);
        $len = count($zpath);
        $mkdir = '';
        for ($i = 0; $i < $len - $oi; $i++) {
            if (!$this->isempt($zpath[$i])) {
                $mkdir .= '/' . $zpath[$i] . '';
                $wzdir = app()->getRootPath() . 'upload' . '' . $mkdir;
                if (!is_dir($wzdir)) {
                    mkdir($wzdir);
                }
            }
        }
    }

    /**
     *    判断变量是否为空
     * @return boolean
     */
    function isempt($str)
    {
        $bool = false;
        if (($str == '' || $str == NULL || empty($str)) && (!is_numeric($str))) $bool = true;
        return $bool;
    }

    public function test()
    {

        $data = [
            [
                ['test' => '测试1', 'aaa' => 'AAA'],
                ['aaa' => 'AAA', 'test' => '测试1']
            ],
            [
                ['test' => '测试2', 'aaa' => 'AAA'],
                ['aaa' => 'AAA', 'test' => '测试2']
            ]
        ];

        $data = [
            ['test' => '测试', 'aaa' => 'AAA'],
            ['aaa' => 'AAA', 'test' => '测试']
        ];
        $this->sheet(2)
            ->filename('测试.xlsx')
            ->sheetname('工作簿名|工作簿2')
            ->title(['标题1','标题2'])
            ->header(['李四:aaa:15:align', '张三:test:16:left'])
            ->border('thin')
            ->body($data)
            ->write();

    }
}