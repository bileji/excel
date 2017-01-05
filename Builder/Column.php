<?php
/**
 * this source file is Column.php
 *
 * author: shuc <shuc324@gmail.com>
 * time:   2016-11-18 10-19
 */

namespace Bileji\Excel\Builder;

use ArrayAccess;

class Column extends Header implements ArrayAccess
{
    private $column;

    public function offsetExists($offset)
    {
        return isset($this->column[$offset]);
    }

    public function offsetGet($offset)
    {
        return $this->column[$offset];
    }

    public function offsetSet($offset, $value)
    {
        $this->column[$offset] = $value;
    }

    public function offsetUnset($offset)
    {
        unset($this->column[$offset]);
    }

    public function toArray()
    {
        return $this->column;
    }

    public function __construct(array $template)
    {
        $this->column = $this->getColumnByTemplate($template);
    }

    private function getColumnByTemplate(array $template, &$column = [], &$position = 'A')
    {
        $stepLength = count($template);
        foreach ($template as $stepName => $item) {
            is_integer($stepName) && $stepName = static::STEP_NAME . ($stepName + 1);
            $column[$stepName] = [
                static::COLUMN_START  => $position,
                static::COLUMN_OVER   => $position,
                static::CHILDREN_LIST => [],
            ];
            if (is_array($item)) {
                $this->getColumnByTemplate($item, $column[$stepName][static::CHILDREN_LIST], $position);
                $column[$stepName][static::COLUMN_OVER] = $position;
            }
            --$stepLength > 0 && ++$position;
        }
        return $column;
    }
}