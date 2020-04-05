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

    /**
     * @var array<string, CharPointer>
     */
    protected array $charPointers = [];

    public function __destruct()
    {
        $this->free();
    }

    public function free(): void
    {
        if (!FFI::isNull($this->pointer)) {
            FFI::free($this->pointer);
        }

        /** @var CharPointer $charPointer */
        foreach ($this->charPointers as $charPointer) {
            $charPointer->free();
        }
        $this->charPointers = [];
    }

    public function getPointer(): CData
    {
        return $this->pointer;
    }

    public function getStruct(): CData
    {
        return $this->struct;
    }

    public function __get($name)
    {
        $getter = 'get' . ucfirst($name) . 'Property';
        if (method_exists($this, $getter)) {
            return $this->{$getter}();
        }

        if ($this->isCharPointerProperty($name)) {
            return $this->getCharPointerPropertyByName($name);
        }

        return $this->struct->{$name};
    }

    public function __set($name, $value)
    {
        $setter = 'set' . ucfirst($name) . 'Property';
        if (method_exists($this, $setter)) {
            $this->{$setter}($value);
            return;
        }

        if ($this->isCharPointerProperty($name)) {
            $this->setCharPointerPropertyByName($name, $value);
            return;
        }

        $this->struct->{$name} = $value;
    }

    /**
     * @return string[] property names
     */
    protected function getCharPointerProperties(): array
    {
        return [];
    }

    /**
     * @param string $name
     * @return bool
     */
    protected function isCharPointerProperty(string $name): bool
    {
        return in_array($name, $this->getCharPointerProperties(), true);
    }

    /**
     * @param string $name
     * @return string
     */
    private function getCharPointerPropertyByName(string $name): ?string
    {
        if (array_key_exists($name, $this->charPointers)) {
            /** @var CharPointer $charPointer */
            $charPointer = $this->charPointers[$name];
            return $charPointer->getString();
        }

        return null;
    }

    /**
     * @param string $name
     * @param string|null $value
     */
    protected function setCharPointerPropertyByName(string $name, ?string $value): void
    {
        if (array_key_exists($name, $this->charPointers)) {
            /** @var CharPointer $charPointer */
            $charPointer = $this->charPointers[$name];
            $charPointer->free();
            unset($this->charPointers[$name]);
        }

        if (null !== $value) {
            $charPointer = new CharPointer($value);
            $this->charPointers[$name] = $charPointer;
            $this->struct->{$name} = $charPointer->getPointer();
        }
    }
}
