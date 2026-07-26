<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Http;

use Mnb\PHPExcel\Support\MnbExcelException;

final class FileRateLimiter implements RateLimiterInterface
{
    public function __construct(private readonly string $directory)
    {
        if (!is_dir($directory) && !mkdir($directory,0775,true) && !is_dir($directory)) { throw new MnbExcelException('Unable to create rate-limit directory.'); }
    }

    public function consume(string $key, int $limit, int $windowSeconds): array
    {
        $limit=max(1,$limit); $windowSeconds=max(1,$windowSeconds); $now=time();
        $path=$this->directory.DIRECTORY_SEPARATOR.hash('sha256',$key).'.json';
        $handle=fopen($path,'c+');
        if ($handle===false) { throw new MnbExcelException('Unable to open rate-limit store.'); }
        try {
            flock($handle,LOCK_EX); $raw=stream_get_contents($handle); $data=json_decode($raw?:'{}',true); $start=(int)($data['start']??$now); $count=(int)($data['count']??0);
            if ($start+$windowSeconds <= $now) { $start=$now; $count=0; }
            $allowed=$count<$limit; if ($allowed) { $count++; }
            ftruncate($handle,0); rewind($handle); fwrite($handle,json_encode(['start'=>$start,'count'=>$count],JSON_THROW_ON_ERROR)); fflush($handle);
            $reset=$start+$windowSeconds;
            return ['allowed'=>$allowed,'limit'=>$limit,'remaining'=>max(0,$limit-$count),'retry_after'=>$allowed?0:max(1,$reset-$now),'reset'=>$reset];
        } finally { flock($handle,LOCK_UN); fclose($handle); }
    }
}
