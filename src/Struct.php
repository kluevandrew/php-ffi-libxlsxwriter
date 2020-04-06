<?php

namespace FFILibXlsxWriter;

use FFI;
use FFI\CData;
use FFILibXlsxWriter\Contracts\DoNotFreeDirectly;

abstract class Struct
{
    /**
     * @var CData|null
     */
    protected ?CData $struct = null;

    /**
     * @var CData|null
     */
    protected ?CData $pointer = null;

    /**
     * @var array<string, Struct>
     */
    protected array $structures = [];

    public function __destruct()
    {
        $this->free();
    }

    protected function canBeFree(): bool
    {
        return $this->pointer !== null
            && !$this instanceof DoNotFreeDirectly
            && !FFI::isNull($this->pointer);
    }

    protected function free(): void
    {
        if ($this->canBeFree()) {
            FFI::free($this->pointer);
        }

        $this->pointer = null;
        $this->struct = null;

        /** @var Struct $structure */
        foreach ($this->structures as $structure) {
            $structure->free();
        }
        $this->structures = [];
    }

    public function getPointer(): ?CData
    {
        return $this->pointer;
    }

    public function getStruct(): ?CData
    {
        return $this->struct;
    }

    public function __get($name)
    {
        $getter = 'get' . ucfirst($name) . 'Property';
        if (method_exists($this, $getter)) {
            return $this->{$getter}();
        }

        if ($this->isStructuralProperty($name)) {
            return $this->getStructuralPropertyByName($name);
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

        if ($this->isStructuralProperty($name)) {
            $this->setStructuralPropertyByName($name, $value);
            return;
        }

        $this->struct->{$name} = $value;
    }

    /**
     * @return string[] property names
     */
    protected function getStructuralProperties(): array
    {
        return [];
    }

    /**
     * @param string $name
     * @return array
     */
    protected function getStructuralPropertyType(string $name): ?array
    {
        $types = $this->getStructuralProperties();
        if (array_key_exists($name, $types)) {
            $definition = $types[$name];
            if (!is_array($definition)) {
                return [$definition, false];
            }

            return $definition;
        }

        throw new \RuntimeException('It\'s not a structural property');
    }

    /**
     * @param string $name
     * @return bool
     */
    protected function isStructuralProperty(string $name): bool
    {
        return array_key_exists($name, $this->getStructuralProperties());
    }

    /**
     * @param string $name
     * @return Struct
     */
    private function getStructuralPropertyByName(string $name): Struct
    {
        if (array_key_exists($name, $this->structures)) {
            return $this->structures[$name];
        }

        return null;
    }

    /**
     * @param string $name
     * @param string|null $value
     */
    protected function setStructuralPropertyByName(string $name, $value): void
    {
        if (array_key_exists($name, $this->structures)) {
            $struct = $this->structures[$name];
            $struct->free();
            unset($this->structures[$name]);
        }

        if (null === $value) {
            return;
        }

        list($class, $pointer) = $this->getStructuralPropertyType($name);

        if (!$value instanceof $class) {
            $value = new $class(...(is_array($value) ? $value : [$value]));
        }

        if (!$value instanceof Struct) {
            throw new \RuntimeException('Should be instance of struct');
        }

        $this->structures[$name] = $value;
        $this->struct->{$name} = $pointer ? $value->getPointer() : $value->getStruct();
    }

    /**
     * @param Struct $struct
     * @return Struct
     */
    protected function addStructure(Struct $struct): Struct
    {
        $this->structures[] = $struct;

        return $struct;
    }

    /**
     * @return Struct[]
     */
    public function getStructures(): array
    {
        return $this->structures;
    }
}
