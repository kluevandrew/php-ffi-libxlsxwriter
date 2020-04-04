<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFI\CData;

abstract class Struct
{
    /**
     * @var CData|mixed
     */
    protected CData $struct;

    /**
     * @var CData
     */
    protected CData $pointer;

    public function __destruct()
    {
        $this->free();
    }

    public function getPointer(): CData
    {
        return $this->pointer;
    }

    public function getStruct(): CData
    {
        return $this->struct;
    }

    public function free(): void
    {
        if (!FFI::isNull($this->pointer)) {
            FFI::free($this->pointer);
        }
    }

    public function __get($name)
    {
        $getter = 'get' . ucfirst($name) . 'Attribute';
        if (method_exists($this, $getter)) {
            return $this->{$getter}();
        }

        return $this->struct->{$name};
    }

    public function __set($name, $value)
    {
        $setter = 'set' . ucfirst($name) . 'Attribute';
        if (method_exists($this, $setter)) {
            $this->{$setter}($value);
            return;
        }

        $this->struct->{$name} = $name;
    }
}
