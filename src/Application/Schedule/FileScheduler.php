<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Schedule;

use DateTimeImmutable;
use DateTimeInterface;
use Mnb\PHPExcel\Support\MnbExcelException;
use Throwable;

final class FileScheduler
{
    public function __construct(private readonly string $path)
    {
        $directory = dirname($path);
        if (!is_dir($directory) && !mkdir($directory, 0775, true) && !is_dir($directory)) {
            throw new MnbExcelException('Unable to create scheduler directory: ' . $directory);
        }
        if (!is_file($path)) {
            $this->save(['tasks' => [], 'runs' => []]);
        }
    }

    /** @param array<string,mixed> $payload */
    public function add(string $id, string $cron, string $type, array $payload, bool $enabled = true): void
    {
        new CronExpression($cron);
        $data = $this->load();
        $data['tasks'][$id] = compact('id', 'cron', 'type', 'payload', 'enabled') + ['updated_at' => gmdate(DATE_ATOM)];
        $this->save($data);
    }

    public function remove(string $id): void
    {
        $data = $this->load();
        unset($data['tasks'][$id], $data['runs'][$id]);
        $this->save($data);
    }

    /** @param callable(string,array<string,mixed>,array<string,mixed>):mixed $dispatcher @return array<string,mixed> */
    public function runDue(callable $dispatcher, ?DateTimeInterface $now = null): array
    {
        $now ??= new DateTimeImmutable('now');
        $minuteKey = $now->format('Y-m-d H:i');
        $data = $this->load();
        $results = [];
        foreach ((array) ($data['tasks'] ?? []) as $id => $task) {
            if (!(bool) ($task['enabled'] ?? true) || (string) ($data['runs'][$id]['minute'] ?? '') === $minuteKey) {
                continue;
            }
            if (!(new CronExpression((string) $task['cron']))->isDue($now)) {
                continue;
            }
            try {
                $result = $dispatcher((string) $task['type'], (array) ($task['payload'] ?? []), $task);
                $results[] = ['id' => $id, 'status' => 'completed', 'result' => $result];
                $data['runs'][$id] = ['minute' => $minuteKey, 'status' => 'completed', 'ran_at' => $now->format(DATE_ATOM)];
            } catch (Throwable $exception) {
                $results[] = ['id' => $id, 'status' => 'failed', 'error' => $exception->getMessage()];
                $data['runs'][$id] = ['minute' => $minuteKey, 'status' => 'failed', 'ran_at' => $now->format(DATE_ATOM), 'error' => $exception->getMessage()];
            }
        }
        $this->save($data);
        return ['status' => 'completed', 'ran' => count($results), 'results' => $results];
    }

    /** @return array<string,mixed> */
    public function all(): array
    {
        return $this->load();
    }

    /** @return array<string,mixed> */
    private function load(): array
    {
        $handle = fopen($this->path, 'c+');
        if ($handle === false) {
            throw new MnbExcelException('Unable to open scheduler store.');
        }
        try {
            flock($handle, LOCK_SH);
            $contents = stream_get_contents($handle);
            $data = json_decode($contents === false || trim($contents) === '' ? '{}' : $contents, true, 512, JSON_THROW_ON_ERROR);
            return is_array($data) ? $data + ['tasks' => [], 'runs' => []] : ['tasks' => [], 'runs' => []];
        } finally {
            flock($handle, LOCK_UN);
            fclose($handle);
        }
    }

    /** @param array<string,mixed> $data */
    private function save(array $data): void
    {
        $tmp = $this->path . '.tmp-' . bin2hex(random_bytes(4));
        $json = json_encode($data, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_THROW_ON_ERROR);
        if (file_put_contents($tmp, $json . "\n", LOCK_EX) === false || !rename($tmp, $this->path)) {
            @unlink($tmp);
            throw new MnbExcelException('Unable to save scheduler store.');
        }
    }
}
