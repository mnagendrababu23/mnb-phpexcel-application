<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Plugin;

interface MnbExcelPluginInterface
{
    public function register(PluginRegistry $registry): void;
}
