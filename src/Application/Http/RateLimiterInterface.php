<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Http;

interface RateLimiterInterface
{
    /** @return array{allowed:bool,limit:int,remaining:int,retry_after:int,reset:int} */
    public function consume(string $key, int $limit, int $windowSeconds): array;
}
