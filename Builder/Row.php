<?php
/**
 * this source file is Row.php
 *
 * author: shuc <shuc324@gmail.com>
 * time:   2016-11-18 18-43
 */

namespace Bileji\Excel\Builder;

use Iterator;
use ArrayAccess;

class Row extends Header implements ArrayAccess, Iterator
{
    private $excelTitle;

    private $maxColumn;

    private $columnBuilder;

    public function offsetUnset($offset)
    {

    }

    public function offsetSet($offset, $value)
    {
    }

    public function offsetGet($offset)
    {
        return $this->excelTitle[$offset];
    }

    public function offsetExists($offset)
    {
        return isset($this->excelTitle[$offset]);
    }

    public function current()
    {
        return current($this->excelTitle);
    }

    public function next()
    {
        return next($this->excelTitle);
    }

    public function rewind()
    {
        return reset($this->excelTitle);
    }

    public function key()
    {
        return key($this->excelTitle);
    }

    public function valid()
    {
        return key($this->excelTitle) !== null;
    }

    public function __construct(Column $columnBuilder)
    {
        $this->columnBuilder = $columnBuilder;
        $this->excelTitle = $this->getExcelTitleByColumnBuilder($this->columnBuilder->toArray());
    }

    private function getExcelTitleByColumnBuilder(&$column, $level = 1, $deep = 1, $top = true)
    {
        static $dimension;
        empty($dimension) && $dimension = $this->getDimension($column);
        if (!isset($column[static::CHILDREN_LIST]) || !empty($column[static::CHILDREN_LIST])) {
            if (isset($column[static::CHILDREN_LIST])) {
                $column[static::RAW_START] = $this->getRawStartNumber($dimension, $deep, $level);
                $column[static::RAW_OVER] = $this->getRawOverNumber($dimension, $deep, $level);
                foreach ($column[static::CHILDREN_LIST] as $title => &$template) {
                    $top === true && $deep = $this->getDimension($template);
                    $this->getExcelTitleByColumnBuilder($template, $level + 1, $deep, false);
                }
            } else {
                foreach ($column as $title => &$template) {
                    $top === true && $deep = $this->getDimension($template);
                    $this->getExcelTitleByColumnBuilder($template, $level + 1, $deep, false);
                }
            }
        } else {
            $column[static::RAW_START] = $this->getRawStartNumber($dimension, $deep, $level);
            $column[static::RAW_OVER] = empty($column[static::CHILDREN_LIST]) ? $dimension : $this->getRawOverNumber($dimension, $deep, $level);
        }
        return $column;
    }

    private function getRawStartNumber($dimension, $deep, $level)
    {
        return ($level - 2) * intval($dimension / $deep) + $level > $dimension % $deep ? 1 : 2;
    }

    private function getRawOverNumber($dimension, $deep, $level)
    {
        return ($level - 1) * intval($dimension / $deep) + $level > $dimension % $deep ? 0 : 1;
    }

    private function getDimension($excelTitle, $dimension = 0, &$maxDimension = 0)
    {
        if (!isset($excelTitle[static::CHILDREN_LIST])) {
            foreach ($excelTitle as $item) {
                $this->getDimension($item, $dimension, $maxDimension);
            }
        } else {
            $maxDimension = ++$dimension > $maxDimension ? $dimension : $maxDimension;
            if (!empty($excelTitle[static::CHILDREN_LIST])) {
                $this->getDimension($excelTitle[static::CHILDREN_LIST], $dimension, $maxDimension);
            }
        }
        return $maxDimension;
    }
}