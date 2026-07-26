<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Http;

use Mnb\PHPExcel\Support\MnbExcelException;
use PDO;

final class PdoRateLimiter implements RateLimiterInterface
{
    public function __construct(private readonly PDO $pdo, private readonly string $table='mnb_excel_rate_limits', bool $autoCreate=true)
    {
        if (preg_match('/^[A-Za-z_][A-Za-z0-9_]*$/',$table)!==1) { throw new MnbExcelException('Invalid rate-limit table name.'); }
        $pdo->setAttribute(PDO::ATTR_ERRMODE,PDO::ERRMODE_EXCEPTION);
        if ($autoCreate) { $pdo->exec("CREATE TABLE IF NOT EXISTS {$table} (key_hash VARCHAR(64) PRIMARY KEY, window_start BIGINT NOT NULL, hit_count INTEGER NOT NULL)"); }
    }
    public function consume(string $key,int $limit,int $windowSeconds):array
    {
        $limit=max(1,$limit);$windowSeconds=max(1,$windowSeconds);$now=time();$hash=hash('sha256',$key);
        $this->pdo->beginTransaction();
        try {
            $stmt=$this->pdo->prepare("SELECT window_start,hit_count FROM {$this->table} WHERE key_hash=:key");$stmt->execute([':key'=>$hash]);$row=$stmt->fetch(PDO::FETCH_ASSOC);
            $start=is_array($row)?(int)$row['window_start']:$now;$count=is_array($row)?(int)$row['hit_count']:0;
            if ($start+$windowSeconds<=$now){$start=$now;$count=0;}
            $allowed=$count<$limit;if($allowed){$count++;}
            $driver=(string)$this->pdo->getAttribute(PDO::ATTR_DRIVER_NAME);
            $sql=$driver==='mysql'?"INSERT INTO {$this->table} (key_hash,window_start,hit_count) VALUES (:key,:start,:count) ON DUPLICATE KEY UPDATE window_start=VALUES(window_start),hit_count=VALUES(hit_count)":"INSERT INTO {$this->table} (key_hash,window_start,hit_count) VALUES (:key,:start,:count) ON CONFLICT(key_hash) DO UPDATE SET window_start=excluded.window_start,hit_count=excluded.hit_count";
            $this->pdo->prepare($sql)->execute([':key'=>$hash,':start'=>$start,':count'=>$count]);$this->pdo->commit();$reset=$start+$windowSeconds;
            return ['allowed'=>$allowed,'limit'=>$limit,'remaining'=>max(0,$limit-$count),'retry_after'=>$allowed?0:max(1,$reset-$now),'reset'=>$reset];
        }catch(\Throwable $e){if($this->pdo->inTransaction())$this->pdo->rollBack();throw $e;}
    }
}
