<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Schedule;

use DateTimeImmutable;
use DateTimeInterface;
use Mnb\PHPExcel\Support\MnbExcelException;
use PDO;
use Throwable;

/** Transactional scheduler store for multi-host deployments. */
final class PdoScheduler implements SchedulerStoreInterface
{
    public function __construct(private readonly PDO $pdo, private readonly string $table = 'mnb_excel_schedule', bool $autoCreate = true)
    {
        if (preg_match('/^[A-Za-z_][A-Za-z0-9_]*$/', $table) !== 1) {
            throw new MnbExcelException('Invalid scheduler table name.');
        }
        $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        if ($autoCreate) { $this->createTable(); }
    }

    public function createTable(): void
    {
        $this->pdo->exec("CREATE TABLE IF NOT EXISTS {$this->table} (id VARCHAR(190) PRIMARY KEY, cron VARCHAR(190) NOT NULL, type VARCHAR(190) NOT NULL, payload TEXT NOT NULL, enabled INTEGER NOT NULL DEFAULT 1, last_minute VARCHAR(16) NULL, last_status VARCHAR(30) NULL, last_run_at VARCHAR(40) NULL, last_result TEXT NULL, last_error TEXT NULL, updated_at VARCHAR(40) NOT NULL)");
    }

    public function add(string $id, string $cron, string $type, array $payload, bool $enabled = true): void
    {
        new CronExpression($cron);
        $params = [':id'=>$id, ':cron'=>$cron, ':type'=>$type, ':payload'=>json_encode($payload, JSON_UNESCAPED_SLASHES|JSON_UNESCAPED_UNICODE|JSON_THROW_ON_ERROR), ':enabled'=>$enabled?1:0, ':updated_at'=>gmdate(DATE_ATOM)];
        $driver = (string) $this->pdo->getAttribute(PDO::ATTR_DRIVER_NAME);
        $sql = $driver === 'mysql'
            ? "INSERT INTO {$this->table} (id,cron,type,payload,enabled,updated_at) VALUES (:id,:cron,:type,:payload,:enabled,:updated_at) ON DUPLICATE KEY UPDATE cron=VALUES(cron),type=VALUES(type),payload=VALUES(payload),enabled=VALUES(enabled),updated_at=VALUES(updated_at)"
            : "INSERT INTO {$this->table} (id,cron,type,payload,enabled,updated_at) VALUES (:id,:cron,:type,:payload,:enabled,:updated_at) ON CONFLICT(id) DO UPDATE SET cron=excluded.cron,type=excluded.type,payload=excluded.payload,enabled=excluded.enabled,updated_at=excluded.updated_at";
        $this->pdo->prepare($sql)->execute($params);
    }

    public function remove(string $id): void
    {
        $this->pdo->prepare("DELETE FROM {$this->table} WHERE id=:id")->execute([':id'=>$id]);
    }

    public function runDue(callable $dispatcher, ?DateTimeInterface $now = null): array
    {
        $now ??= new DateTimeImmutable('now');
        $minute = $now->format('Y-m-d H:i');
        $rows = $this->pdo->query("SELECT * FROM {$this->table} WHERE enabled=1 ORDER BY id")->fetchAll(PDO::FETCH_ASSOC);
        $results=[];
        foreach ($rows as $task) {
            if ((string)($task['last_minute']??'')===$minute || !(new CronExpression((string)$task['cron']))->isDue($now)) { continue; }
            $claim=$this->pdo->prepare("UPDATE {$this->table} SET last_minute=:minute_set,last_status='running',last_run_at=:ran WHERE id=:id AND (last_minute IS NULL OR last_minute<>:minute_check)");
            $claim->execute([':minute_set'=>$minute, ':minute_check'=>$minute, ':ran'=>$now->format(DATE_ATOM), ':id'=>$task['id']]);
            if ($claim->rowCount()!==1) { continue; }
            try {
                $payload=json_decode((string)$task['payload'],true,512,JSON_THROW_ON_ERROR);
                $result=$dispatcher((string)$task['type'], is_array($payload)?$payload:[], $task);
                $this->pdo->prepare("UPDATE {$this->table} SET last_status='completed',last_result=:result,last_error=NULL WHERE id=:id")->execute([':result'=>json_encode($result,JSON_UNESCAPED_SLASHES|JSON_UNESCAPED_UNICODE|JSON_THROW_ON_ERROR),':id'=>$task['id']]);
                $results[]=['id'=>$task['id'],'status'=>'completed','result'=>$result];
            } catch (Throwable $e) {
                $this->pdo->prepare("UPDATE {$this->table} SET last_status='failed',last_error=:error WHERE id=:id")->execute([':error'=>$e->getMessage(),':id'=>$task['id']]);
                $results[]=['id'=>$task['id'],'status'=>'failed','error'=>$e->getMessage()];
            }
        }
        return ['status'=>'completed','ran'=>count($results),'results'=>$results];
    }

    public function all(): array
    {
        $rows=$this->pdo->query("SELECT * FROM {$this->table} ORDER BY id")->fetchAll(PDO::FETCH_ASSOC);
        return ['tasks'=>$rows];
    }
}
