<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Queue;

use Mnb\PHPExcel\Support\MnbExcelException;
use PDO;
use PDOException;

/**
 * Transactional queue backend suitable for multiple hosts and workers.
 * Supports SQLite, MySQL/MariaDB and PostgreSQL through portable PDO SQL.
 */
final class PdoQueue implements QueueBackendInterface
{
    private string $driver;

    public function __construct(private readonly PDO $pdo, private readonly string $table = 'mnb_excel_queue', bool $autoCreate = true)
    {
        if (preg_match('/^[A-Za-z_][A-Za-z0-9_]*$/', $table) !== 1) {
            throw new MnbExcelException('Invalid queue table name.');
        }
        $this->driver = (string) $pdo->getAttribute(PDO::ATTR_DRIVER_NAME);
        $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        if ($autoCreate) {
            $this->createTable();
        }
    }

    public function createTable(): void
    {
        $idType = $this->driver === 'pgsql' ? 'VARCHAR(80)' : 'VARCHAR(80)';
        $textType = $this->driver === 'pgsql' ? 'TEXT' : 'TEXT';
        $sql = "CREATE TABLE IF NOT EXISTS {$this->table} (\n"
            . "id {$idType} PRIMARY KEY, type VARCHAR(190) NOT NULL, payload {$textType} NOT NULL, state VARCHAR(20) NOT NULL,\n"
            . "attempts INTEGER NOT NULL DEFAULT 0, available_at BIGINT NOT NULL, reserved_at BIGINT NULL, reserved_by VARCHAR(190) NULL,\n"
            . "created_at VARCHAR(40) NOT NULL, completed_at VARCHAR(40) NULL, failed_at VARCHAR(40) NULL, result {$textType} NULL,\n"
            . "last_error {$textType} NULL, exception VARCHAR(255) NULL)";
        $this->pdo->exec($sql);
        foreach (["CREATE INDEX IF NOT EXISTS {$this->table}_ready_idx ON {$this->table} (state, available_at)", "CREATE INDEX IF NOT EXISTS {$this->table}_reserved_idx ON {$this->table} (state, reserved_at)"] as $indexSql) {
            try { $this->pdo->exec($indexSql); } catch (PDOException) { /* Some older MySQL versions lack IF NOT EXISTS for indexes. */ }
        }
    }

    public function enqueue(string $type, array $payload, int $delaySeconds = 0): QueueJob
    {
        $type = trim($type);
        if ($type === '') {
            throw new MnbExcelException('Queue job type cannot be empty.');
        }
        $id = gmdate('YmdHis') . '-' . bin2hex(random_bytes(8));
        $job = new QueueJob($id, $type, $payload, 0, time() + max(0, $delaySeconds), gmdate(DATE_ATOM));
        $stmt = $this->pdo->prepare("INSERT INTO {$this->table} (id,type,payload,state,attempts,available_at,created_at) VALUES (:id,:type,:payload,'pending',0,:available_at,:created_at)");
        $stmt->execute([
            ':id' => $job->id,
            ':type' => $job->type,
            ':payload' => json_encode($payload, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_THROW_ON_ERROR),
            ':available_at' => $job->availableAt,
            ':created_at' => $job->createdAt,
        ]);
        return $job;
    }

    public function reserve(int $visibilityTimeoutSeconds = 300, ?string $workerId = null): ?QueueJob
    {
        $this->releaseExpired($visibilityTimeoutSeconds);
        $workerId ??= (gethostname() ?: 'worker') . ':' . getmypid() . ':' . bin2hex(random_bytes(3));
        $now = time();
        for ($attempt = 0; $attempt < 5; $attempt++) {
            $this->pdo->beginTransaction();
            try {
                $suffix = in_array($this->driver, ['mysql', 'pgsql'], true) ? ' FOR UPDATE SKIP LOCKED' : '';
                $stmt = $this->pdo->prepare("SELECT id,type,payload,attempts,available_at,created_at FROM {$this->table} WHERE state='pending' AND available_at<=:now ORDER BY available_at,id LIMIT 1" . $suffix);
                $stmt->execute([':now' => $now]);
                $row = $stmt->fetch(PDO::FETCH_ASSOC);
                if (!is_array($row)) {
                    $this->pdo->commit();
                    return null;
                }
                $update = $this->pdo->prepare("UPDATE {$this->table} SET state='processing', attempts=attempts+1, reserved_at=:reserved_at, reserved_by=:reserved_by WHERE id=:id AND state='pending'");
                $update->execute([':reserved_at' => $now, ':reserved_by' => $workerId, ':id' => $row['id']]);
                if ($update->rowCount() !== 1) {
                    $this->pdo->rollBack();
                    continue;
                }
                $this->pdo->commit();
                return new QueueJob(
                    (string) $row['id'],
                    (string) $row['type'],
                    $this->decode((string) $row['payload']),
                    (int) $row['attempts'] + 1,
                    (int) $row['available_at'],
                    (string) $row['created_at']
                );
            } catch (\Throwable $e) {
                if ($this->pdo->inTransaction()) { $this->pdo->rollBack(); }
                throw $e;
            }
        }
        return null;
    }

    public function complete(QueueJob $job, array $result = []): void
    {
        $stmt = $this->pdo->prepare("UPDATE {$this->table} SET state='completed', completed_at=:at, result=:result, reserved_at=NULL, reserved_by=NULL WHERE id=:id AND state='processing'");
        $stmt->execute([':at' => gmdate(DATE_ATOM), ':result' => json_encode($result, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_THROW_ON_ERROR), ':id' => $job->id]);
    }

    public function fail(QueueJob $job, \Throwable $exception, int $retryDelaySeconds = 0, int $maxAttempts = 3): void
    {
        if ($job->attempts < max(1, $maxAttempts)) {
            $stmt = $this->pdo->prepare("UPDATE {$this->table} SET state='pending', available_at=:available_at, reserved_at=NULL, reserved_by=NULL, last_error=:error, exception=:exception WHERE id=:id");
            $stmt->execute([':available_at' => time() + max(0, $retryDelaySeconds), ':error' => $exception->getMessage(), ':exception' => $exception::class, ':id' => $job->id]);
            return;
        }
        $stmt = $this->pdo->prepare("UPDATE {$this->table} SET state='failed', failed_at=:at, reserved_at=NULL, reserved_by=NULL, last_error=:error, exception=:exception WHERE id=:id");
        $stmt->execute([':at' => gmdate(DATE_ATOM), ':error' => $exception->getMessage(), ':exception' => $exception::class, ':id' => $job->id]);
    }

    public function releaseExpired(int $visibilityTimeoutSeconds = 300): int
    {
        $stmt = $this->pdo->prepare("UPDATE {$this->table} SET state='pending', available_at=:now, reserved_at=NULL, reserved_by=NULL, last_error='Job reservation expired and was released.' WHERE state='processing' AND reserved_at IS NOT NULL AND reserved_at<=:cutoff");
        $stmt->execute([':now' => time(), ':cutoff' => time() - max(1, $visibilityTimeoutSeconds)]);
        return $stmt->rowCount();
    }

    public function stats(): array
    {
        $stats = ['pending' => 0, 'processing' => 0, 'completed' => 0, 'failed' => 0];
        $stmt = $this->pdo->query("SELECT state, COUNT(*) AS total FROM {$this->table} GROUP BY state");
        foreach ($stmt->fetchAll(PDO::FETCH_ASSOC) as $row) {
            if (isset($stats[$row['state']])) { $stats[$row['state']] = (int) $row['total']; }
        }
        return $stats;
    }

    /** @return array<string,mixed> */
    private function decode(string $json): array
    {
        $value = json_decode($json, true, 512, JSON_THROW_ON_ERROR);
        return is_array($value) ? $value : [];
    }
}
