<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

/**
 * Lightweight package-readiness checks for the public Composer release.
 * This intentionally avoids network access and does not require vendor/.
 */
final class ReleaseReadiness
{
    /** @return array{status:string,checks:list<array{name:string,status:string,message:string}>,summary:array{passed:int,warning:int,failed:int}} */
    public static function check(string $root): array
    {
        $root = rtrim($root, DIRECTORY_SEPARATOR);
        $checks = [];

        foreach (['composer.json', 'README.md', 'CHANGELOG.md', 'ROADMAP.md', 'LICENSE', 'src', 'examples', 'tests', '.gitignore'] as $path) {
            $checks[] = self::checkExists($root, $path);
        }

        $checks[] = self::checkComposer($root . DIRECTORY_SEPARATOR . 'composer.json');
        $checks[] = self::checkNoVendor($root);
        $checks[] = self::checkNoGitDirectory($root);
        $checks[] = self::checkNoComposerLock($root);
        $checks[] = self::checkNoGeneratedOutputs($root);
        $checks[] = self::checkPhpFilesSyntaxReady($root);

        $summary = ['passed' => 0, 'warning' => 0, 'failed' => 0];
        foreach ($checks as $check) {
            if ($check['status'] === 'pass') {
                $summary['passed']++;
            } elseif ($check['status'] === 'warning') {
                $summary['warning']++;
            } else {
                $summary['failed']++;
            }
        }

        return [
            'status' => $summary['failed'] > 0 ? 'failed' : ($summary['warning'] > 0 ? 'warning' : 'ready'),
            'checks' => $checks,
            'summary' => $summary,
        ];
    }

    /** @return array{name:string,status:string,message:string} */
    private static function checkExists(string $root, string $path): array
    {
        $exists = file_exists($root . DIRECTORY_SEPARATOR . $path);
        return [
            'name' => 'exists:' . $path,
            'status' => $exists ? 'pass' : 'fail',
            'message' => $exists ? $path . ' exists.' : $path . ' is missing.',
        ];
    }

    /** @return array{name:string,status:string,message:string} */
    private static function checkComposer(string $path): array
    {
        if (!is_file($path)) {
            return ['name' => 'composer', 'status' => 'fail', 'message' => 'composer.json missing.'];
        }

        $json = json_decode((string) file_get_contents($path), true);
        if (!is_array($json)) {
            return ['name' => 'composer', 'status' => 'fail', 'message' => 'composer.json is invalid JSON.'];
        }

        $missing = [];
        foreach (['name', 'description', 'type', 'license', 'autoload', 'require'] as $key) {
            if (!array_key_exists($key, $json)) {
                $missing[] = $key;
            }
        }

        if ($missing !== []) {
            return ['name' => 'composer', 'status' => 'fail', 'message' => 'composer.json missing keys: ' . implode(', ', $missing) . '.'];
        }

        $autoload = $json['autoload']['psr-4']['Mnb\\PHPExcel\\'] ?? null;
        if ($autoload !== 'src/') {
            return ['name' => 'composer', 'status' => 'fail', 'message' => 'PSR-4 autoload must map Mnb\\PHPExcel\\ to src/.'];
        }

        return ['name' => 'composer', 'status' => 'pass', 'message' => 'composer.json metadata and autoload look ready.'];
    }

    /** @return array{name:string,status:string,message:string} */
    private static function checkNoVendor(string $root): array
    {
        $exists = is_dir($root . DIRECTORY_SEPARATOR . 'vendor');
        return [
            'name' => 'no-vendor',
            'status' => $exists ? 'warning' : 'pass',
            'message' => $exists ? 'vendor/ exists in this working directory. It should not be included in a clean release ZIP.' : 'vendor/ is not included.',
        ];
    }


    /** @return array{name:string,status:string,message:string} */
    private static function checkNoGitDirectory(string $root): array
    {
        $exists = is_dir($root . DIRECTORY_SEPARATOR . '.git');
        return [
            'name' => 'no-git-directory',
            'status' => $exists ? 'warning' : 'pass',
            'message' => $exists ? '.git/ exists in this working directory. It should not be included in a clean public release ZIP.' : '.git/ is not included.',
        ];
    }

    /** @return array{name:string,status:string,message:string} */
    private static function checkNoComposerLock(string $root): array
    {
        $exists = is_file($root . DIRECTORY_SEPARATOR . 'composer.lock');
        return [
            'name' => 'no-composer-lock',
            'status' => $exists ? 'warning' : 'pass',
            'message' => $exists ? 'composer.lock is usually not included in a reusable Composer library package.' : 'composer.lock is not included.',
        ];
    }

    /** @return array{name:string,status:string,message:string} */
    private static function checkNoGeneratedOutputs(string $root): array
    {
        $bad = [];
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($root, \FilesystemIterator::SKIP_DOTS));
        foreach ($iterator as $file) {
            /** @var \SplFileInfo $file */
            if (!$file->isFile()) {
                continue;
            }
            $relative = substr($file->getPathname(), strlen($root) + 1);
            if (str_starts_with($relative, 'examples' . DIRECTORY_SEPARATOR . 'assets' . DIRECTORY_SEPARATOR)) {
                continue;
            }
            $extension = strtolower($file->getExtension());
            if (in_array($extension, ['xlsx', 'xls', 'csv'], true)) {
                $bad[] = $relative;
            }
        }

        return [
            'name' => 'no-generated-outputs',
            'status' => $bad === [] ? 'pass' : 'warning',
            'message' => $bad === [] ? 'No generated XLSX/CSV files found.' : 'Generated sample output files found: ' . implode(', ', array_slice($bad, 0, 8)),
        ];
    }

    /** @return array{name:string,status:string,message:string} */
    private static function checkPhpFilesSyntaxReady(string $root): array
    {
        $src = $root . DIRECTORY_SEPARATOR . 'src';
        if (!is_dir($src)) {
            return ['name' => 'php-files', 'status' => 'fail', 'message' => 'src/ is missing.'];
        }

        $count = 0;
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($src, \FilesystemIterator::SKIP_DOTS));
        foreach ($iterator as $file) {
            /** @var \SplFileInfo $file */
            if ($file->isFile() && $file->getExtension() === 'php') {
                $count++;
            }
        }

        return ['name' => 'php-files', 'status' => $count > 0 ? 'pass' : 'fail', 'message' => $count . ' PHP source files found.'];
    }
}
