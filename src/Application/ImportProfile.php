<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application;

use Mnb\PHPExcel\MnbExcel;
use PDO;

final class ImportProfile
{
    private ?string $sourcePath = null;

    /** @param array<string,mixed> $profile */
    public function __construct(private string $name, private array $profile)
    {
    }

    public function source(string $path): self
    {
        $clone = clone $this;
        $clone->sourcePath = $path;
        return $clone;
    }

    /** @return array<string,mixed> */
    public function config(): array
    {
        return $this->profile;
    }

    /** @return array<string,mixed> */
    public function plan(array $serverOptions = []): array
    {
        $path = $this->sourcePath ?? (string) ($this->profile['source_path'] ?? '');
        return MnbExcel::autoImportPlan($path, $serverOptions, is_array($this->profile['analysis_options'] ?? null) ? $this->profile['analysis_options'] : []);
    }

    /** @param PDO|array<string,mixed>|string|null $pdo @param array<string,mixed> $overrides @return array<string,mixed> */
    public function run(PDO|array|string|null $pdo = null, array $overrides = []): array
    {
        $path = $this->sourcePath ?? (string) ($this->profile['source_path'] ?? '');
        $table = (string) ($overrides['table'] ?? $this->profile['table'] ?? '');
        $db = $pdo ?? ($overrides['db'] ?? $this->profile['db'] ?? null);
        $options = array_merge($this->profile, $overrides, ['profile' => $this->name]);
        unset($options['source_path'], $options['table'], $options['db']);
        return MnbExcel::largeImportToSql($path, $db, $table, $options);
    }
}
