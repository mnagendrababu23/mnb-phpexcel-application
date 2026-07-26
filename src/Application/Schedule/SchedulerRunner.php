<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Schedule;

use Mnb\PHPExcel\Support\MnbExcelException;

/** Long-running scheduler runner with process lock and graceful signals. */
final class SchedulerRunner
{
    private bool $running = true;

    public function __construct(private readonly SpreadsheetScheduler $scheduler)
    {
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function runForever(array $options = []): array
    {
        $interval = max(1, (int)($options['interval_seconds'] ?? 30));
        $maxCycles = max(0, (int)($options['max_cycles'] ?? 0));
        $lockPath = (string)($options['lock_path'] ?? (sys_get_temp_dir().DIRECTORY_SEPARATOR.'mnb-phpexcel-scheduler.lock'));
        $lock = fopen($lockPath, 'c+');
        if ($lock === false || !flock($lock, LOCK_EX | LOCK_NB)) {
            if (is_resource($lock)) { fclose($lock); }
            throw new MnbExcelException('Another spreadsheet scheduler runner already holds the lock: '.$lockPath);
        }
        if (function_exists('pcntl_async_signals')) {
            pcntl_async_signals(true);
            foreach ([SIGTERM, SIGINT, SIGHUP] as $signal) { pcntl_signal($signal, fn()=> $this->running=false); }
        }
        $cycles=0; $ran=0; $failures=0; $started=microtime(true); $results=[];
        try {
            while ($this->running && ($maxCycles===0 || $cycles<$maxCycles)) {
                $cycles++;
                $result=$this->scheduler->runDue();
                $ran += (int)($result['ran']??0);
                foreach ((array)($result['results']??[]) as $item) { if (($item['status']??'')==='failed') { $failures++; } }
                if (($options['collect_results'] ?? false) === true) { $results[]=$result; }
                if (!$this->running || ($maxCycles>0 && $cycles >= $maxCycles)) { break; }
                sleep($interval);
            }
        } finally {
            flock($lock, LOCK_UN); fclose($lock);
            if (($options['remove_lock_file'] ?? false) === true) { @unlink($lockPath); }
        }
        return ['status'=>$failures===0?'completed':'completed_with_errors','cycles'=>$cycles,'jobs_run'=>$ran,'failures'=>$failures,'elapsed_seconds'=>round(microtime(true)-$started,6),'results'=>$results];
    }

    public function stop(): void { $this->running=false; }
}
