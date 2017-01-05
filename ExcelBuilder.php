<?php
/**
 * this source file is ExcelBuilder.php
 *
 * author: shuc <shuc324@gmail.com>
 * time:   2016-11-16 15-57
 */

namespace Bileji\Excel;

use Bileji\Excel\Builder\Column;
use Bileji\Excel\Builder\Header;
use Bileji\Excel\Builder\Row;

class ExcelBuilder extends Header
{

    /**
     * 将数组关系转换为点关系列表
     * @param array $template 数据模板
     * @param array $keyDotList 点关系列表
     * @param null $parentKeyDot 父级点关系
     * @return array 点关系列表
     */
    private function exchangeArrayMapToDotRelationList(array $template, &$keyDotList = [], $parentKeyDot = null)
    {
        foreach ($template as $key => $value) {
            $key = is_null($parentKeyDot) ? $key : ($parentKeyDot . static::DELIMITER . (string)$key);
            !empty($value[static::CHILDREN_LIST]) ? $this->exchangeArrayMapToDotRelationList($value[static::CHILDREN_LIST], $key, $keyDotList) : ($keyDotList[] = $key);
        }
        return $keyDotList;
    }

    /**
     * 通过点关系从对象中取值
     * @param array $data 数据
     * @param string $dotRelation 点关系
     * @return mixed
     */
    private function getValueByDotRelation(array $data, $dotRelation)
    {
        array_map(function ($key) use (&$data) {
            $key = preg_match('/^\d+$/', $key) ? (int)$key : $key;
            $data = isset($data[$key]) ? $data[$key] : null;
        }, explode(static::DELIMITER, $dotRelation));
        return $data;
    }

    /*** todo ***/
    private $data;

    private $PHPExcel;
    private $filename;

    private $maxColumn = [];

    public function maxRaw()
    {
        foreach ($this->data as $column) {
            foreach ($column as $name => $item) {
                if (!is_array($item)) {
                    empty($this->maxColumn[$name]) && $this->maxColumn[$name] = $item;
                } else {
                    empty($this->maxColumn[$name]) && $this->maxColumn[$name] = [];
                    $this->maxColumn[$name] += $item;
                }
            }
        }
        return $this;
    }

    public function __construct(PHPExcel $PHPExcel, array $data, $filename = '')
    {
        $this->maxColumn = empty($data[0]) ? [] : $data[0];
        $this->PHPExcel = $PHPExcel;
        $this->data = $data;
        $this->filename = $filename;
    }

    public function create()
    {
        return $this->createTitleByTemplate()->addContent()->saveExcel();
    }

    private function createTitleByTemplate()
    {
        $template = new Row(new Column($this->maxColumn));
        $this->createExcelTitleRecursive($this->PHPExcel, $template);
        return $this;
    }

    private function createExcelTitleRecursive(PHPExcel $PHPExcel, $template, $title = '')
    {
        $activeSheet = $PHPExcel->getActiveSheet();
        if (isset($template[static::CHILDREN_LIST])) {
            $activeSheet->setCellValue($template[static::COLUMN_START] . $template[static::RAW_START], $title);
            if ($template[static::COLUMN_START] != $template[static::COLUMN_OVER] || $template[static::RAW_START] != $template[static::RAW_OVER]) {
                $activeSheet->mergeCells($template[static::COLUMN_START] . $template[static::RAW_START] . ':' . $template[static::COLUMN_OVER] . $template[static::RAW_OVER]);
            }
            $alignment = $activeSheet->getStyle($template[static::COLUMN_START] . $template[static::RAW_START])->getAlignment();
            $alignment->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $alignment->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            if (!empty($template[static::CHILDREN_LIST])) {
                foreach ($template[static::CHILDREN_LIST] as $title => $item) {
                    $this->createExcelTitleRecursive($PHPExcel, $item, $title);
                }
            }
        } else {
            foreach ($template as $title => $item) {
                $this->createExcelTitleRecursive($PHPExcel, $item, $title);
            }
        }
    }

    private function addContent()
    {

        return $this;
    }

    private function saveExcel()
    {
        $this->PHPExcel->getActiveSheet()->setAutoFilter($this->PHPExcel->getActiveSheet()->calculateWorksheetDimension());
        PHPExcel_IOFactory::createWriter($this->PHPExcel, 'Excel2007')->save($this->filename);
        return $this->filename;
    }
}