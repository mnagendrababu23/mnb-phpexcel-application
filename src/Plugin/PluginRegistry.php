<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Plugin;

use Mnb\PHPExcel\Application\ImportProfileManager;
use Mnb\PHPExcel\Application\RowTransformerPipeline;
use Mnb\PHPExcel\Events\EventDispatcher;
use Mnb\PHPExcel\Validation\CustomValidatorRegistry;
use Mnb\PHPExcel\Contracts\ReaderPluginInterface;
use Mnb\PHPExcel\MnbExcel;

final class PluginRegistry
{
    /** @param array<string,mixed> $profile */
    public function importProfile(string $name, array $profile): self
    {
        ImportProfileManager::register($name, $profile);
        return $this;
    }

    /** @param callable(array<string|int,mixed>, array<string,mixed>):array<string|int,mixed> $transformer */
    public function transformer(string $name, callable $transformer): self
    {
        RowTransformerPipeline::register($name, $transformer);
        return $this;
    }

    /** @param callable(mixed,array<string,mixed>,array<string,mixed>):bool|string|null $callback */
    public function validator(string $name, callable $callback): self
    {
        CustomValidatorRegistry::register($name, $callback);
        return $this;
    }

    public function reader(ReaderPluginInterface $plugin, int $priority = 0): self
    {
        MnbExcel::registerReaderPlugin($plugin, $priority);
        return $this;
    }

    /** @param callable(array<string,mixed>):mixed $listener */
    public function event(string $event, callable $listener): self
    {
        EventDispatcher::listen($event, $listener);
        return $this;
    }
}
