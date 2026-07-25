<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Plugin;

final class PluginManager
{
    /** @var list<string> */
    private static array $plugins = [];

    public static function register(MnbExcelPluginInterface|callable $plugin): void
    {
        $registry = new PluginRegistry();
        if ($plugin instanceof MnbExcelPluginInterface) {
            $plugin->register($registry);
            self::$plugins[] = $plugin::class;
            return;
        }
        $plugin($registry);
        self::$plugins[] = 'closure@' . count(self::$plugins);
    }

    /** @return list<string> */
    public static function plugins(): array
    {
        return self::$plugins;
    }

    public static function clear(): void
    {
        self::$plugins = [];
    }
}
