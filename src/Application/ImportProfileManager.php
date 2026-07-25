<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

final class ImportProfileManager
{
    /** @var array<string,array<string,mixed>> */
    private static array $profiles = [];

    /** @param array<string,mixed>|string $profiles */
    public static function load(array|string $profiles): void
    {
        if (is_string($profiles)) {
            if (!is_file($profiles)) {
                throw MnbExcelException::withCode('Import profile file not found: ' . $profiles, ErrorCode::FILE_NOT_FOUND);
            }
            $loaded = require $profiles;
            if (!is_array($loaded)) {
                throw MnbExcelException::withCode('Import profile file must return an array.', ErrorCode::VALIDATION_FAILED);
            }
            $profiles = $loaded;
        }
        foreach ($profiles as $name => $profile) {
            if (is_array($profile)) {
                self::register((string) $name, $profile);
            }
        }
    }

    /** @param array<string,mixed> $profile */
    public static function register(string $name, array $profile): void
    {
        self::$profiles[$name] = $profile;
    }

    /** @return array<string,mixed> */
    public static function get(string $name): array
    {
        if (!isset(self::$profiles[$name])) {
            throw MnbExcelException::withCode('Import profile not found: ' . $name, ErrorCode::VALIDATION_FAILED);
        }
        return self::$profiles[$name];
    }

    public static function profile(string $name): ImportProfile
    {
        return new ImportProfile($name, self::get($name));
    }

    /** @return array<string,array<string,mixed>> */
    public static function all(): array
    {
        return self::$profiles;
    }

    public static function clear(): void
    {
        self::$profiles = [];
    }
}
