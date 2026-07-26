<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Queue;

use Mnb\PHPExcel\Support\MnbExcelException;

final class FileQueue implements QueueBackendInterface
{
    public function __construct(private readonly string $directory)
    {
        foreach (['pending', 'processing', 'completed', 'failed'] as $name) {
            $path = $this->directory . DIRECTORY_SEPARATOR . $name;
            if (!is_dir($path) && !mkdir($path, 0775, true) && !is_dir($path)) {
                throw new MnbExcelException('Unable to create queue directory: ' . $path);
            }
        }
    }

    /** @param array<string,mixed> $payload */
    public function enqueue(string $type, array $payload, int $delaySeconds = 0): QueueJob
    {
        $type = trim($type);
        if ($type === '') {
            throw new MnbExcelException('Queue job type cannot be empty.');
        }
        $id = gmdate('YmdHis') . '-' . bin2hex(random_bytes(8));
        $job = new QueueJob($id, $type, $payload, 0, time() + max(0, $delaySeconds), gmdate(DATE_ATOM));
        $this->write($this->path('pending', $id), $job->toArray());
        return $job;
    }

    public function reserve(int $visibilityTimeoutSeconds = 300, ?string $workerId = null): ?QueueJob
    {
        $this->releaseExpired($visibilityTimeoutSeconds);
        $files = glob($this->directory . DIRECTORY_SEPARATOR . 'pending' . DIRECTORY_SEPARATOR . '*.json') ?: [];
        sort($files, SORT_STRING);
        foreach ($files as $file) {
            $data = $this->read($file);
            if ((int) ($data['available_at'] ?? 0) > time()) {
                continue;
            }
            $id = (string) ($data['id'] ?? pathinfo($file, PATHINFO_FILENAME));
            $target = $this->path('processing', $id);
            if (!@rename($file, $target)) {
                continue;
            }
            $data['attempts'] = (int) ($data['attempts'] ?? 0) + 1;
            $data['reserved_at'] = time();
            $data['reserved_by'] = $workerId ?? (gethostname() . ':' . getmypid());
            $this->write($target, $data);
            return QueueJob::fromArray($data);
        }
        return null;
    }

    /** @param array<string,mixed> $result */
    public function complete(QueueJob $job, array $result = []): void
    {
        $source = $this->path('processing', $job->id);
        $data = $job->toArray() + ['completed_at' => gmdate(DATE_ATOM), 'result' => $result];
        $this->write($this->path('completed', $job->id), $data);
        @unlink($source);
    }

    public function fail(QueueJob $job, \Throwable $exception, int $retryDelaySeconds = 0, int $maxAttempts = 3): void
    {
        $source = $this->path('processing', $job->id);
        if ($job->attempts < max(1, $maxAttempts)) {
            $data = $job->toArray();
            $data['available_at'] = time() + max(0, $retryDelaySeconds);
            $data['last_error'] = $exception->getMessage();
            $this->write($this->path('pending', $job->id), $data);
        } else {
            $data = $job->toArray() + [
                'failed_at' => gmdate(DATE_ATOM),
                'error' => $exception->getMessage(),
                'exception' => $exception::class,
            ];
            $this->write($this->path('failed', $job->id), $data);
        }
        @unlink($source);
    }


    public function releaseExpired(int $visibilityTimeoutSeconds = 300): int
    {
        $released = 0;
        $cutoff = time() - max(1, $visibilityTimeoutSeconds);
        $files = glob($this->directory . DIRECTORY_SEPARATOR . 'processing' . DIRECTORY_SEPARATOR . '*.json') ?: [];
        foreach ($files as $file) {
            $data = $this->read($file);
            $reservedAt = (int) ($data['reserved_at'] ?? filemtime($file) ?: 0);
            if ($reservedAt > $cutoff) {
                continue;
            }
            $id = (string) ($data['id'] ?? pathinfo($file, PATHINFO_FILENAME));
            $data['available_at'] = time();
            $data['last_error'] = 'Job reservation expired and was released.';
            unset($data['reserved_at'], $data['reserved_by']);
            $target = $this->path('pending', $id);
            $this->write($target, $data);
            if (@unlink($file)) {
                $released++;
            }
        }
        return $released;
    }

    /** @return array<string,int> */
    public function stats(): array
    {
        $stats = [];
        foreach (['pending', 'processing', 'completed', 'failed'] as $name) {
            $stats[$name] = count(glob($this->directory . DIRECTORY_SEPARATOR . $name . DIRECTORY_SEPARATOR . '*.json') ?: []);
        }
        return $stats;
    }

    private function path(string $state, string $id): string
    {
        return $this->directory . DIRECTORY_SEPARATOR . $state . DIRECTORY_SEPARATOR . basename($id) . '.json';
    }

    /** @param array<string,mixed> $data */
    private function write(string $path, array $data): void
    {
        $json = json_encode($data, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_THROW_ON_ERROR);
        $tmp = $path . '.tmp-' . bin2hex(random_bytes(4));
        if (file_put_contents($tmp, $json . "\n", LOCK_EX) === false || !rename($tmp, $path)) {
            @unlink($tmp);
            throw new MnbExcelException('Unable to write queue job: ' . $path);
        }
    }

    /** @return array<string,mixed> */
    private function read(string $path): array
    {
        $data = json_decode((string) file_get_contents($path), true, 512, JSON_THROW_ON_ERROR);
        return is_array($data) ? $data : [];
    }
}
